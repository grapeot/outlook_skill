from __future__ import annotations

from datetime import datetime
from typing import cast

import httpx
from markdown import markdown as md_to_html

from .auth import AuthManager
from .config import Settings
from .errors import GraphApiError, OutlookSkillError


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
    dry_run: bool = False,
) -> dict[str, object]:
    if not subject.strip():
        raise OutlookSkillError("calendar invite requires --subject.")
    if not attendees and not optional_attendees:
        raise OutlookSkillError("calendar invite requires at least one attendee.")
    if body_format not in ("text", "html", "markdown", "md"):
        raise OutlookSkillError(f"Unsupported body format: {body_format}")
    _validate_time_range(start, end)

    content_type, content_value = _prepare_body(body_text, body_format)
    graph_payload: dict[str, object] = {
        "subject": subject,
        "body": {"contentType": content_type, "content": content_value},
        "start": {"dateTime": start, "timeZone": timezone},
        "end": {"dateTime": end, "timeZone": timezone},
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

    result = {
        "dry_run": dry_run,
        "endpoint": "/me/calendar/events",
        "subject": subject,
        "start": start,
        "end": end,
        "timezone": timezone,
        "attendees": list(attendees),
        "optional_attendees": list(optional_attendees),
        "location": location,
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


def _validate_time_range(start: str, end: str) -> None:
    try:
        start_dt = _parse_iso_like(start)
        end_dt = _parse_iso_like(end)
    except ValueError as exc:
        raise OutlookSkillError("calendar invite requires ISO-like --start and --end values.") from exc
    if end_dt <= start_dt:
        raise OutlookSkillError("calendar invite requires --end to be after --start.")


def _parse_iso_like(value: str) -> datetime:
    return datetime.fromisoformat(value.replace("Z", "+00:00"))


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
        "$select": "subject,start,end,recurrence,importance,isAllDay,showAs,location,webLink",
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
