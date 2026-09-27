from __future__ import annotations

from datetime import datetime
from typing import cast
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

import httpx
from markdown import markdown as md_to_html

from .auth import AuthManager
from .config import Settings
from .errors import GraphApiError, OutlookSkillError

_TIMEZONE_ALIASES: dict[str, str] = {
    "PT": "America/Los_Angeles",
    "PST": "America/Los_Angeles",
    "PDT": "America/Los_Angeles",
    "PACIFIC": "America/Los_Angeles",
    "US/PACIFIC": "America/Los_Angeles",
    "PACIFIC STANDARD TIME": "America/Los_Angeles",
    "MT": "America/Denver",
    "CT": "America/Chicago",
    "ET": "America/New_York",
}


def resolve_timezone(name: str) -> str:
    key = name.strip()
    if key.upper() == "UTC":
        return "UTC"
    if key.upper() in _TIMEZONE_ALIASES:
        return _TIMEZONE_ALIASES[key.upper()]
    if not key:
        raise OutlookSkillError("calendar invite requires a non-empty --timezone value.")
    try:
        ZoneInfo(key)
    except (ZoneInfoNotFoundError, ValueError) as exc:
        aliases = ", ".join(sorted(set(_TIMEZONE_ALIASES) | {"UTC"}))
        raise OutlookSkillError(
            f"calendar invite --timezone {name!r} is not a recognized timezone. "
            f"Use an IANA zone name (e.g. America/Los_Angeles) or an alias: {aliases}."
        ) from exc
    return key


def create_calendar_invite(
    settings: Settings,
    *,
    subject: str,
    start: str,
    end: str,
    timezone: str = "UTC",
    attendees: tuple[str, ...],
    body_text: str = "",
    body_format: str = "text",
    location: str | None = None,
    optional_attendees: tuple[str, ...] = (),
    reminder_minutes: int | None = None,
    dry_run: bool = False,
) -> dict[str, object]:
    if not subject.strip():
        raise OutlookSkillError("calendar invite requires --subject.")
    if body_format not in ("text", "html", "markdown", "md"):
        raise OutlookSkillError(f"Unsupported body format: {body_format}")
    if reminder_minutes is not None and reminder_minutes < 0:
        raise OutlookSkillError("calendar invite requires --reminder-minutes to be zero or positive.")
    resolved_timezone = resolve_timezone(timezone)
    zone = ZoneInfo(resolved_timezone)
    start_dt = _parse_wall_clock(start, "--start")
    end_dt = _parse_wall_clock(end, "--end")
    if end_dt <= start_dt:
        raise OutlookSkillError("calendar invite requires --end to be after --start.")
    start_aware = _to_aware_local(start_dt, zone, "--start")
    end_aware = _to_aware_local(end_dt, zone, "--end")
    _validate_absolute_order(start_aware, end_aware)

    content_type, content_value = _prepare_body(body_text, body_format)
    graph_payload: dict[str, object] = {
        "subject": subject,
        "body": {"contentType": content_type, "content": content_value},
        "start": {"dateTime": start, "timeZone": resolved_timezone},
        "end": {"dateTime": end, "timeZone": resolved_timezone},
        "attendees": [
            {**_recipient(addr), "type": "required"}
            for addr in attendees
        ] + [
            {**_recipient(addr), "type": "optional"}
            for addr in optional_attendees
        ],
    }
    if location:
        graph_payload["location"] = {"displayName": location}
    if reminder_minutes is not None:
        graph_payload["reminderMinutesBeforeStart"] = reminder_minutes

    result = {
        "dry_run": dry_run,
        "endpoint": "/me/calendar/events",
        "subject": subject,
        "start": start,
        "end": end,
        "timezone": resolved_timezone,
        "start_utc": _to_utc_iso(start_aware),
        "end_utc": _to_utc_iso(end_aware),
        "attendees": list(attendees),
        "optional_attendees": list(optional_attendees),
        "location": location,
        "reminder_minutes": reminder_minutes,
        "body_content_type": content_type,
        "body_chars": len(content_value),
    }
    if dry_run:
        return {**result, "created": False, "note": "dry run; Graph /me/calendar/events was not called"}

    token = AuthManager(settings).get_access_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Content-Type": "application/json",
    }
    with httpx.Client(base_url=settings.graph_base_url, timeout=httpx.Timeout(120.0, connect=10.0), headers=headers) as client:
        response = client.post("/me/calendar/events", json=graph_payload)
    if response.is_error:
        raise GraphApiError(
            f"Graph calendar event creation failed with status {response.status_code}",
            status_code=response.status_code,
            response_text=response.text,
        )
    payload = cast(dict[str, object], response.json() if response.content else {})
    event_id = payload.get("id")
    web_link = payload.get("webLink")
    return {
        **result,
        "created": True,
        "event_id": event_id if isinstance(event_id, str) else None,
        "web_link": web_link if isinstance(web_link, str) else None,
    }


def _prepare_body(body_text: str, body_format: str) -> tuple[str, str]:
    if body_format in ("markdown", "md"):
        return "HTML", md_to_html(body_text, extensions=["extra", "sane_lists"])
    if body_format == "html":
        return "HTML", body_text
    return "Text", body_text


def _recipient(address: str) -> dict[str, dict[str, str]]:
    return {"emailAddress": {"address": address}}


def _parse_wall_clock(value: str, flag: str) -> datetime:
    try:
        parsed = datetime.fromisoformat(value)
    except ValueError as exc:
        raise OutlookSkillError(
            f"calendar invite requires ISO-like {flag} values, e.g. 2026-05-06T10:00:00."
        ) from exc
    if parsed.tzinfo is not None:
        raise OutlookSkillError(
            f"calendar invite {flag} must be a wall-clock time without a UTC offset "
            f"(e.g. 2026-10-01T18:00:00); the zone is set by --timezone."
        )
    return parsed


def _validate_absolute_order(start_aware: datetime, end_aware: datetime) -> None:
    # Compare as absolute instants: aware datetimes sharing one ZoneInfo
    # object compare by wall time in CPython, which inverts across DST gaps.
    if end_aware.astimezone(ZoneInfo("UTC")) <= start_aware.astimezone(ZoneInfo("UTC")):
        raise OutlookSkillError("calendar invite requires --end to be after --start in absolute time.")


def _to_aware_local(dt: datetime, zone: ZoneInfo, flag: str) -> datetime:
    aware = dt.replace(tzinfo=zone)
    round_trip = aware.astimezone(ZoneInfo("UTC")).astimezone(zone).replace(tzinfo=None)
    if round_trip != dt:
        raise OutlookSkillError(
            f"calendar invite {flag} {dt.isoformat()} does not exist in {zone.key} "
            f"(DST spring-forward gap). Use a valid wall-clock time in that zone."
        )
    # Ambiguous fall-back times keep Python's fold=0 default: first occurrence
    # (DST offset, the earlier UTC instant).
    return aware


def _to_utc_iso(aware: datetime) -> str:
    return aware.astimezone(ZoneInfo("UTC")).isoformat().replace("+00:00", "Z")


# --- calendar list ---


def list_calendar_events(
    settings: Settings,
    *,
    start_date: str,
    end_date: str,
    skip_recurring: str = "daily,weekly",
) -> dict[str, object]:
    token = AuthManager(settings).get_access_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    }

    start_dt = f"{start_date}T00:00:00" if "T" not in start_date else start_date
    end_dt = f"{end_date}T23:59:59" if "T" not in end_date else end_date

    params: dict[str, str | int] = {
        "$filter": f"start/dateTime ge '{start_dt}' and start/dateTime le '{end_dt}'",
        "$orderby": "start/dateTime",
        "$select": "id,subject,start,end,recurrence,importance,isAllDay,showAs,location,webLink",
        "$top": 100,
    }

    with httpx.Client(base_url=settings.graph_base_url, timeout=httpx.Timeout(120.0, connect=10.0), headers=headers) as client:
        response = client.get("/me/calendar/events", params=params)

    if response.is_error:
        raise GraphApiError(
            f"Graph calendar list failed with status {response.status_code}",
            status_code=response.status_code,
            response_text=response.text,
        )

    payload = cast(dict[str, object], response.json() if response.content else {})
    raw_events = cast(list[dict[str, object]], payload.get("value", [])) if isinstance(payload.get("value"), list) else []

    total_count = len(raw_events)
    filtered_events: list[dict[str, object]] = []
    skipped_recurring = 0

    skip_types: set[str] | None = None  # None = "none", empty = "all", specific = matched types
    if skip_recurring == "all":
        skip_types = set()
    elif skip_recurring and skip_recurring != "none":
        skip_types = set(t.strip() for t in skip_recurring.split(",") if t.strip())

    for raw in raw_events:
        recurrence = raw.get("recurrence")
        rec_type = _extract_recurrence_type(recurrence) if isinstance(recurrence, dict) else None

        should_skip = False
        if rec_type is not None:
            if skip_types is None:
                should_skip = False
            elif len(skip_types) == 0:
                should_skip = True
            elif rec_type in skip_types:
                should_skip = True

        if should_skip:
            skipped_recurring += 1
            continue

        start_data = raw.get("start")
        start_dt_str = ""
        tz = ""
        if isinstance(start_data, dict):
            start_dt_str = str(start_data.get("dateTime", ""))
            tz = str(start_data.get("timeZone", ""))

        end_data = raw.get("end")
        end_dt_str = ""
        if isinstance(end_data, dict):
            end_dt_str = str(end_data.get("dateTime", ""))

        loc_data = raw.get("location")
        loc_name = ""
        if isinstance(loc_data, dict):
            loc_name = str(loc_data.get("displayName", ""))

        filtered_events.append({
            "event_id": str(raw.get("id", "")),
            "subject": str(raw.get("subject", "")),
            "start": start_dt_str,
            "end": end_dt_str,
            "timezone": tz,
            "is_all_day": bool(raw.get("isAllDay", False)),
            "importance": str(raw.get("importance", "normal")),
            "show_as": str(raw.get("showAs", "free")),
            "location": loc_name,
            "recurrence_type": rec_type,
            "web_link": str(raw.get("webLink", "")),
        })

    return {
        "start_date": start_date,
        "end_date": end_date,
        "total_count": total_count,
        "shown_count": len(filtered_events),
        "skipped_recurring": skipped_recurring,
        "events": filtered_events,
    }


def _extract_recurrence_type(recurrence: dict[str, object]) -> str | None:
    pattern = recurrence.get("pattern")
    if isinstance(pattern, dict):
        rec_type = pattern.get("type")
        if isinstance(rec_type, str):
            return rec_type
    return None


def delete_calendar_event(
    settings: Settings,
    *,
    event_id: str,
    dry_run: bool = False,
) -> dict[str, object]:
    if not event_id.strip():
        raise OutlookSkillError("calendar delete requires --event-id.")

    result = {
        "dry_run": dry_run,
        "endpoint": f"/me/calendar/events/{event_id}",
        "event_id": event_id,
    }
    if dry_run:
        return {**result, "deleted": False, "note": "dry run; Graph DELETE was not called"}

    token = AuthManager(settings).get_access_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    }
    with httpx.Client(base_url=settings.graph_base_url, timeout=httpx.Timeout(120.0, connect=10.0), headers=headers) as client:
        get_response = client.get(f"/me/calendar/events/{event_id}")
        if get_response.is_error:
            raise GraphApiError(
                f"Graph calendar event fetch before deletion failed with status {get_response.status_code}",
                status_code=get_response.status_code,
                response_text=get_response.text,
            )
        deleted_event = cast(dict[str, object], get_response.json() if get_response.content else {})
        response = client.delete(f"/me/calendar/events/{event_id}")
    if response.is_error:
        raise GraphApiError(
            f"Graph calendar event deletion failed with status {response.status_code}",
            status_code=response.status_code,
            response_text=response.text,
        )
    return {**result, "deleted": True, "deleted_event": deleted_event}


def get_calendar_event(
    settings: Settings,
    *,
    event_id: str,
) -> dict[str, object]:
    if not event_id.strip():
        raise OutlookSkillError("calendar get requires --event-id.")

    token = AuthManager(settings).get_access_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    }
    with httpx.Client(base_url=settings.graph_base_url, timeout=httpx.Timeout(120.0, connect=10.0), headers=headers) as client:
        response = client.get(f"/me/calendar/events/{event_id}")
    if response.is_error:
        raise GraphApiError(
            f"Graph calendar event fetch failed with status {response.status_code}",
            status_code=response.status_code,
            response_text=response.text,
        )
    event = cast(dict[str, object], response.json() if response.content else {})
    return {
        "endpoint": f"/me/calendar/events/{event_id}",
        "event_id": event_id,
        "event": event,
    }
