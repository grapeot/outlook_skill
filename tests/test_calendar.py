from __future__ import annotations

import json
from datetime import datetime
from typing import Any, cast
from urllib.parse import unquote
from zoneinfo import ZoneInfo

import httpx
import pytest

from outlook_skill import calendar
from outlook_skill.config import Settings
from outlook_skill.errors import OutlookSkillError


class FakeTransport(httpx.BaseTransport):
    def __init__(self) -> None:
        self.requests: list[tuple[str, str, Any]] = []

    def handle_request(self, request: httpx.Request) -> httpx.Response:
        self.requests.append((request.method, str(request.url), request.content))
        if request.method == "POST" and str(request.url).endswith("/me/calendar/events"):
            return httpx.Response(201, json={"id": "EVENT_123", "webLink": "https://calendar.example/event"})
        return httpx.Response(404, json={"error": "unexpected"})


def install_fake_graph(monkeypatch, transport: FakeTransport) -> None:
    class FakeAuthManager:
        def __init__(self, settings): pass
        def get_access_token(self): return "TOKEN"

    original_client = httpx.Client

    def fake_client(*args, **kwargs):
        kwargs.pop("transport", None)
        return original_client(*args, transport=transport, **kwargs)

    monkeypatch.setattr(calendar, "AuthManager", FakeAuthManager)
    monkeypatch.setattr(calendar.httpx, "Client", fake_client)


def settings() -> Settings:
    value = object.__new__(type("S", (), {"graph_base_url": "https://example.test/v1.0"}))
    return cast(Settings, cast(object, value))


def test_create_calendar_invite_dry_run_does_not_call_graph(monkeypatch):
    transport = FakeTransport()
    install_fake_graph(monkeypatch, transport)

    payload = calendar.create_calendar_invite(
        settings(),
        subject="Meeting",
        start="2026-05-06T10:00:00",
        end="2026-05-06T10:30:00",
        attendees=("duck@example.com",),
        dry_run=True,
    )

    assert payload["dry_run"] is True
    assert payload["created"] is False
    assert transport.requests == []


def test_create_calendar_invite_posts_to_me_calendar_events(monkeypatch):
    transport = FakeTransport()
    install_fake_graph(monkeypatch, transport)

    payload = calendar.create_calendar_invite(
        settings(),
        subject="Meeting",
        start="2026-05-06T10:00:00",
        end="2026-05-06T10:30:00",
        timezone="Pacific Standard Time",
        attendees=("required@example.com",),
        optional_attendees=("optional@example.com",),
        location="Zoom",
        body_text="Agenda",
    )

    assert payload["created"] is True
    assert payload["event_id"] == "EVENT_123"
    assert payload["timezone"] == "America/Los_Angeles"
    assert payload["start_utc"] == "2026-05-06T17:00:00Z"
    assert payload["end_utc"] == "2026-05-06T17:30:00Z"
    assert transport.requests[0][0] == "POST"
    assert transport.requests[0][1].endswith("/me/calendar/events")
    graph_payload = json.loads(transport.requests[0][2])
    assert graph_payload["subject"] == "Meeting"
    assert graph_payload["start"] == {"dateTime": "2026-05-06T10:00:00", "timeZone": "America/Los_Angeles"}
    assert graph_payload["end"] == {"dateTime": "2026-05-06T10:30:00", "timeZone": "America/Los_Angeles"}
    assert graph_payload["location"] == {"displayName": "Zoom"}
    assert graph_payload["attendees"][0]["emailAddress"]["address"] == "required@example.com"
    assert graph_payload["attendees"][0]["type"] == "required"
    assert graph_payload["attendees"][1]["emailAddress"]["address"] == "optional@example.com"
    assert graph_payload["attendees"][1]["type"] == "optional"


def test_create_calendar_invite_markdown_body_converts_to_html(monkeypatch):
    transport = FakeTransport()
    install_fake_graph(monkeypatch, transport)

    calendar.create_calendar_invite(
        settings(),
        subject="Meeting",
        start="2026-05-06T10:00:00",
        end="2026-05-06T10:30:00",
        attendees=("duck@example.com",),
        body_text="**bold**",
        body_format="markdown",
    )

    graph_payload = json.loads(transport.requests[0][2])
    assert graph_payload["body"]["contentType"] == "HTML"
    assert "<strong>bold</strong>" in graph_payload["body"]["content"]


def test_create_calendar_invite_with_no_attendees(monkeypatch):
    transport = FakeTransport()
    install_fake_graph(monkeypatch, transport)

    payload = calendar.create_calendar_invite(
        settings(),
        subject="Appointment",
        start="2026-05-06T10:00:00",
        end="2026-05-06T10:30:00",
        attendees=(),
    )

    assert payload["created"] is True
    assert payload["attendees"] == []
    graph_payload = json.loads(transport.requests[0][2])
    assert graph_payload["attendees"] == []
    assert "attendees" in graph_payload


def test_create_calendar_invite_rejects_missing_subject():
    with pytest.raises(OutlookSkillError):
        calendar.create_calendar_invite(
            settings(),
            subject=" ",
            start="2026-05-06T10:00:00",
            end="2026-05-06T10:30:00",
            attendees=("duck@example.com",),
        )


def test_create_calendar_invite_rejects_bad_body_format():
    with pytest.raises(OutlookSkillError):
        calendar.create_calendar_invite(
            settings(),
            subject="Meeting",
            start="2026-05-06T10:00:00",
            end="2026-05-06T10:30:00",
            attendees=("duck@example.com",),
            body_format="xml",
        )


def test_create_calendar_invite_rejects_end_before_start():
    with pytest.raises(OutlookSkillError):
        calendar.create_calendar_invite(
            settings(),
            subject="Meeting",
            start="2026-05-06T10:30:00",
            end="2026-05-06T10:00:00",
            attendees=("duck@example.com",),
        )


def test_create_calendar_invite_resolves_pt_alias_with_dst(monkeypatch):
    transport = FakeTransport()
    install_fake_graph(monkeypatch, transport)

    payload = calendar.create_calendar_invite(
        settings(),
        subject="Dinner",
        start="2026-10-01T18:00:00",
        end="2026-10-01T20:00:00",
        timezone="PT",
        attendees=(),
    )

    # October 1 is PDT (UTC-7): 18:00 local is 01:00 UTC the next day.
    assert payload["timezone"] == "America/Los_Angeles"
    assert payload["start_utc"] == "2026-10-02T01:00:00Z"
    assert payload["end_utc"] == "2026-10-02T03:00:00Z"
    graph_payload = json.loads(transport.requests[0][2])
    assert graph_payload["start"] == {"dateTime": "2026-10-01T18:00:00", "timeZone": "America/Los_Angeles"}
    assert graph_payload["end"] == {"dateTime": "2026-10-01T20:00:00", "timeZone": "America/Los_Angeles"}


def test_create_calendar_invite_pt_alias_uses_standard_time_in_winter(monkeypatch):
    transport = FakeTransport()
    install_fake_graph(monkeypatch, transport)

    payload = calendar.create_calendar_invite(
        settings(),
        subject="Dinner",
        start="2026-01-15T18:00:00",
        end="2026-01-15T20:00:00",
        timezone="pt",
        attendees=(),
    )

    # January is PST (UTC-8): 18:00 local is 02:00 UTC the next day.
    assert payload["timezone"] == "America/Los_Angeles"
    assert payload["start_utc"] == "2026-01-16T02:00:00Z"
    assert payload["end_utc"] == "2026-01-16T04:00:00Z"


def test_create_calendar_invite_utc_is_identity(monkeypatch):
    transport = FakeTransport()
    install_fake_graph(monkeypatch, transport)

    payload = calendar.create_calendar_invite(
        settings(),
        subject="Dinner",
        start="2026-10-02T01:00:00",
        end="2026-10-02T03:00:00",
        timezone="UTC",
        attendees=(),
    )

    assert payload["timezone"] == "UTC"
    assert payload["start_utc"] == "2026-10-02T01:00:00Z"
    assert payload["end_utc"] == "2026-10-02T03:00:00Z"
    graph_payload = json.loads(transport.requests[0][2])
    assert graph_payload["start"]["timeZone"] == "UTC"


@pytest.mark.parametrize("bad_start", ["2026-10-01T18:00:00+00:00", "2026-10-01T18:00:00Z"])
def test_create_calendar_invite_rejects_offset_bearing_start(bad_start):
    with pytest.raises(OutlookSkillError, match="wall-clock"):
        calendar.create_calendar_invite(
            settings(),
            subject="Dinner",
            start=bad_start,
            end="2026-10-01T20:00:00",
            timezone="PT",
            attendees=(),
        )


def test_create_calendar_invite_rejects_offset_bearing_end():
    with pytest.raises(OutlookSkillError, match="wall-clock"):
        calendar.create_calendar_invite(
            settings(),
            subject="Dinner",
            start="2026-10-01T18:00:00",
            end="2026-10-01T20:00:00Z",
            timezone="PT",
            attendees=(),
        )


def test_create_calendar_invite_rejects_unknown_timezone():
    with pytest.raises(OutlookSkillError, match="not a recognized timezone"):
        calendar.create_calendar_invite(
            settings(),
            subject="Dinner",
            start="2026-10-01T18:00:00",
            end="2026-10-01T20:00:00",
            timezone="Mars/Olympus_Mons",
            attendees=(),
        )


def test_create_calendar_invite_rejects_nonexistent_spring_forward_time():
    # 2026-03-08 02:00-03:00 PT does not exist (DST spring-forward).
    # Without validation, 02:30 (PST offset) would map to 10:30Z while 03:00
    # (PDT offset) maps to 10:00Z — end before start.
    with pytest.raises(OutlookSkillError, match="does not exist"):
        calendar.create_calendar_invite(
            settings(),
            subject="Dinner",
            start="2026-03-08T02:30:00",
            end="2026-03-08T03:00:00",
            timezone="PT",
            attendees=(),
        )


def test_create_calendar_invite_ambiguous_fall_back_time_uses_first_occurrence(monkeypatch):
    transport = FakeTransport()
    install_fake_graph(monkeypatch, transport)

    payload = calendar.create_calendar_invite(
        settings(),
        subject="Dinner",
        start="2026-11-01T01:30:00",
        end="2026-11-01T02:00:00",
        timezone="PT",
        attendees=(),
    )

    # 01:30 happens twice on 2026-11-01; fold=0 picks the first occurrence (PDT, UTC-7).
    assert payload["start_utc"] == "2026-11-01T08:30:00Z"
    assert payload["end_utc"] == "2026-11-01T10:00:00Z"


def test_create_calendar_invite_preserves_fractional_seconds(monkeypatch):
    transport = FakeTransport()
    install_fake_graph(monkeypatch, transport)

    payload = calendar.create_calendar_invite(
        settings(),
        subject="Dinner",
        start="2026-10-01T18:00:00.500000",
        end="2026-10-01T20:00:00.500000",
        timezone="PT",
        attendees=(),
    )

    assert payload["start_utc"] == "2026-10-02T01:00:00.500000Z"
    assert payload["end_utc"] == "2026-10-02T03:00:00.500000Z"


def test_validate_absolute_order_catches_wall_time_inversion():
    la = ZoneInfo("America/Los_Angeles")
    # Wall time 02:30 < 03:00, but absolute 10:30Z > 10:00Z (gap-time offsets).
    # A shared-ZoneInfo wall-time comparison would pass this pair.
    start = datetime(2026, 3, 8, 2, 30, tzinfo=la)
    end = datetime(2026, 3, 8, 3, 0, tzinfo=la)
    with pytest.raises(OutlookSkillError, match="absolute time"):
        calendar._validate_absolute_order(start, end)


def test_create_calendar_invite_normalizes_lowercase_utc(monkeypatch):
    transport = FakeTransport()
    install_fake_graph(monkeypatch, transport)

    payload = calendar.create_calendar_invite(
        settings(),
        subject="Dinner",
        start="2026-10-02T01:00:00",
        end="2026-10-02T03:00:00",
        timezone="utc",
        attendees=(),
    )

    assert payload["timezone"] == "UTC"
    graph_payload = json.loads(transport.requests[0][2])
    assert graph_payload["start"]["timeZone"] == "UTC"


# --- calendar list tests ---


def sample_graph_event(**overrides: Any) -> dict[str, Any]:
    base = {
        "id": "EVT_001",
        "subject": "Test Event",
        "start": {"dateTime": "2026-05-15T10:00:00.0000000", "timeZone": "Pacific Standard Time"},
        "end": {"dateTime": "2026-05-15T11:00:00.0000000", "timeZone": "Pacific Standard Time"},
        "recurrence": None,
        "importance": "normal",
        "isAllDay": False,
        "showAs": "busy",
        "location": {"displayName": "Room A"},
        "webLink": "https://outlook.live.com/calendar/0/event/EVM_001",
    }
    base.update(overrides)
    return base


class FakeCalendarTransport(httpx.BaseTransport):
    def __init__(self) -> None:
        self.requests: list[tuple[str, str, Any]] = []

    def handle_request(self, request: httpx.Request) -> httpx.Response:
        self.requests.append((request.method, str(request.url), request.content))
        return httpx.Response(200, json={
            "value": [
                sample_graph_event(id="EVT_001", subject="Non-recurring meeting"),
                sample_graph_event(
                    id="EVT_002",
                    subject="Daily standup",
                    recurrence={
                        "pattern": {"type": "daily", "interval": 1},
                        "range": {"type": "endDate", "startDate": "2026-01-01"},
                    },
                ),
                sample_graph_event(
                    id="EVT_003",
                    subject="Weekly sync",
                    recurrence={
                        "pattern": {"type": "weekly", "interval": 1, "daysOfWeek": ["Monday"]},
                        "range": {"type": "endDate", "startDate": "2026-01-01"},
                    },
                ),
                sample_graph_event(
                    id="EVT_004",
                    subject="Monthly review",
                    recurrence={
                        "pattern": {"type": "absoluteMonthly", "interval": 1, "dayOfMonth": 15},
                        "range": {"type": "noEnd", "startDate": "2026-01-01"},
                    },
                ),
                sample_graph_event(
                    id="EVT_005",
                    subject="Yearly planning",
                    recurrence={
                        "pattern": {"type": "absoluteYearly", "interval": 1, "month": 12},
                        "range": {"type": "noEnd", "startDate": "2026-01-01"},
                    },
                ),
            ],
            "@odata.nextLink": None,
        })


def install_fake_calendar_graph(monkeypatch, transport: httpx.BaseTransport) -> None:
    class FakeAuthManager:
        def __init__(self, settings): pass
        def get_access_token(self): return "TOKEN"

    original_client = httpx.Client

    def fake_client(*args, **kwargs):
        kwargs.pop("transport", None)
        return original_client(*args, transport=transport, **kwargs)

    monkeypatch.setattr(calendar, "AuthManager", FakeAuthManager)
    monkeypatch.setattr(calendar.httpx, "Client", fake_client)


def test_list_calendar_events_default_skips_daily_weekly(monkeypatch):
    transport = FakeCalendarTransport()
    install_fake_calendar_graph(monkeypatch, transport)

    result = calendar.list_calendar_events(
        settings(),
        start_date="2026-05-08",
        end_date="2026-07-08",
    )

    subjects = [e["subject"] for e in result["events"]]
    assert "Non-recurring meeting" in subjects
    assert "Monthly review" in subjects
    assert "Yearly planning" in subjects
    assert "Daily standup" not in subjects
    assert "Weekly sync" not in subjects
    assert result["total_count"] == 5
    assert result["shown_count"] == 3
    assert result["skipped_recurring"] == 2
    first = cast(list[dict[str, Any]], result["events"])[0]
    assert first["event_id"] == "EVT_001"


def test_list_calendar_events_skip_recurring_all(monkeypatch):
    transport = FakeCalendarTransport()
    install_fake_calendar_graph(monkeypatch, transport)

    result = calendar.list_calendar_events(
        settings(),
        start_date="2026-05-08",
        end_date="2026-07-08",
        skip_recurring="all",
    )

    subjects = [e["subject"] for e in result["events"]]
    assert "Non-recurring meeting" in subjects
    assert "Monthly review" not in subjects
    assert "Yearly planning" not in subjects
    assert result["shown_count"] == 1


def test_list_calendar_events_skip_recurring_none(monkeypatch):
    transport = FakeCalendarTransport()
    install_fake_calendar_graph(monkeypatch, transport)

    result = calendar.list_calendar_events(
        settings(),
        start_date="2026-05-08",
        end_date="2026-07-08",
        skip_recurring="none",
    )

    assert result["shown_count"] == 5
    assert result["skipped_recurring"] == 0


def test_list_calendar_events_empty_result(monkeypatch):
    class EmptyTransport(httpx.BaseTransport):
        def handle_request(self, request: httpx.Request) -> httpx.Response:
            return httpx.Response(200, json={"value": []})

    install_fake_calendar_graph(monkeypatch, EmptyTransport())

    result = calendar.list_calendar_events(
        settings(),
        start_date="2026-05-08",
        end_date="2026-05-09",
    )

    assert result["total_count"] == 0
    assert result["shown_count"] == 0
    assert result["events"] == []


def test_list_calendar_events_constructs_correct_query(monkeypatch):
    transport = FakeCalendarTransport()
    install_fake_calendar_graph(monkeypatch, transport)

    calendar.list_calendar_events(
        settings(),
        start_date="2026-05-08",
        end_date="2026-07-08",
    )

    assert len(transport.requests) == 1
    method, url, _ = transport.requests[0]
    assert method == "GET"
    assert "/me/calendar/events" in url
    decoded = unquote(url)
    assert "$filter=start/dateTime" in decoded
    assert "2026-05-08" in decoded
    assert "$orderby=start/dateTime" in decoded
    assert "$select=" in decoded
    assert "$top=" in decoded


def test_list_calendar_events_handles_custom_recurring_types(monkeypatch):
    transport = FakeCalendarTransport()
    install_fake_calendar_graph(monkeypatch, transport)

    result = calendar.list_calendar_events(
        settings(),
        start_date="2026-05-08",
        end_date="2026-07-08",
        skip_recurring="daily,weekly,absoluteMonthly",
    )

    subjects = [e["subject"] for e in result["events"]]
    assert "Non-recurring meeting" in subjects
    assert "Yearly planning" in subjects
    assert "Monthly review" not in subjects
    assert result["shown_count"] == 2
    assert result["skipped_recurring"] == 3


class FakeEventDetailTransport(httpx.BaseTransport):
    def __init__(self) -> None:
        self.requests: list[tuple[str, str, Any]] = []

    def handle_request(self, request: httpx.Request) -> httpx.Response:
        self.requests.append((request.method, str(request.url), request.content))
        if request.method == "GET" and str(request.url).endswith("/me/calendar/events/EVT_123"):
            return httpx.Response(200, json={
                "id": "EVT_123",
                "subject": "Placeholder",
                "body": {"contentType": "html", "content": "<p>full body</p>"},
                "attendees": [],
            })
        if request.method == "DELETE" and str(request.url).endswith("/me/calendar/events/EVT_123"):
            return httpx.Response(204)
        return httpx.Response(404, json={"error": "unexpected"})


def test_get_calendar_event_returns_raw_event(monkeypatch):
    transport = FakeEventDetailTransport()
    install_fake_calendar_graph(monkeypatch, transport)

    result = calendar.get_calendar_event(settings(), event_id="EVT_123")

    assert result["event_id"] == "EVT_123"
    assert cast(dict[str, Any], result["event"])["subject"] == "Placeholder"
    assert transport.requests == [("GET", "https://example.test/v1.0/me/calendar/events/EVT_123", b"")]


def test_delete_calendar_event_fetches_full_event_before_delete(monkeypatch):
    transport = FakeEventDetailTransport()
    install_fake_calendar_graph(monkeypatch, transport)

    result = calendar.delete_calendar_event(settings(), event_id="EVT_123")

    assert result["deleted"] is True
    deleted_event = cast(dict[str, Any], result["deleted_event"])
    assert deleted_event["id"] == "EVT_123"
    assert deleted_event["body"] == {"contentType": "html", "content": "<p>full body</p>"}
    assert [request[0] for request in transport.requests] == ["GET", "DELETE"]


def test_delete_calendar_event_dry_run_does_not_call_graph(monkeypatch):
    transport = FakeEventDetailTransport()
    install_fake_calendar_graph(monkeypatch, transport)

    result = calendar.delete_calendar_event(settings(), event_id="EVT_123", dry_run=True)

    assert result["deleted"] is False
    assert transport.requests == []


# --- calendar update tests ---


class FakeUpdateTransport(httpx.BaseTransport):
    def __init__(self) -> None:
        self.requests: list[tuple[str, str, Any]] = []

    def handle_request(self, request: httpx.Request) -> httpx.Response:
        self.requests.append((request.method, str(request.url), request.content))
        if request.method == "PATCH" and str(request.url).endswith("/me/calendar/events/EVT_123"):
            return httpx.Response(200, json={
                "id": "EVT_123",
                "subject": "Updated",
                "webLink": "https://calendar.example/event",
            })
        return httpx.Response(404, json={"error": "unexpected"})


def test_update_calendar_event_dry_run_does_not_call_graph(monkeypatch):
    transport = FakeUpdateTransport()
    install_fake_calendar_graph(monkeypatch, transport)

    result = calendar.update_calendar_event(
        settings(),
        event_id="EVT_123",
        start="2026-04-15T13:00:00",
        end="2026-04-15T14:30:00",
        timezone="PT",
        dry_run=True,
    )

    assert result["dry_run"] is True
    assert result["updated"] is False
    assert transport.requests == []


def test_update_calendar_event_patches_only_provided_fields(monkeypatch):
    transport = FakeUpdateTransport()
    install_fake_calendar_graph(monkeypatch, transport)

    result = calendar.update_calendar_event(
        settings(),
        event_id="EVT_123",
        subject="Updated",
    )

    assert result["updated"] is True
    assert result["changed_fields"] == ["--subject"]
    assert transport.requests[0][0] == "PATCH"
    assert transport.requests[0][1].endswith("/me/calendar/events/EVT_123")
    graph_payload = json.loads(transport.requests[0][2])
    assert graph_payload == {"subject": "Updated"}


def test_update_calendar_event_resolves_timezone_and_echoes_utc(monkeypatch):
    transport = FakeUpdateTransport()
    install_fake_calendar_graph(monkeypatch, transport)

    result = calendar.update_calendar_event(
        settings(),
        event_id="EVT_123",
        start="2026-04-15T13:00:00",
        end="2026-04-15T14:30:00",
        timezone="PT",
    )

    # April 15 is PDT (UTC-7): 13:00 local is 20:00 UTC.
    assert result["timezone"] == "America/Los_Angeles"
    assert result["start_utc"] == "2026-04-15T20:00:00Z"
    assert result["end_utc"] == "2026-04-15T21:30:00Z"
    assert result["changed_fields"] == ["--start/--end"]
    graph_payload = json.loads(transport.requests[0][2])
    assert graph_payload["start"] == {"dateTime": "2026-04-15T13:00:00", "timeZone": "America/Los_Angeles"}
    assert graph_payload["end"] == {"dateTime": "2026-04-15T14:30:00", "timeZone": "America/Los_Angeles"}
    assert "subject" not in graph_payload


def test_update_calendar_event_requires_at_least_one_field():
    with pytest.raises(OutlookSkillError, match="at least one field"):
        calendar.update_calendar_event(settings(), event_id="EVT_123")


def test_update_calendar_event_requires_both_start_and_end():
    with pytest.raises(OutlookSkillError, match="both --start and --end"):
        calendar.update_calendar_event(settings(), event_id="EVT_123", start="2026-04-15T13:00:00")


def test_update_calendar_event_rejects_missing_event_id():
    with pytest.raises(OutlookSkillError, match="--event-id"):
        calendar.update_calendar_event(settings(), event_id=" ", subject="Updated")


def test_update_calendar_event_rejects_end_before_start():
    with pytest.raises(OutlookSkillError, match="--end to be after --start"):
        calendar.update_calendar_event(
            settings(),
            event_id="EVT_123",
            start="2026-04-15T14:30:00",
            end="2026-04-15T13:00:00",
            timezone="PT",
        )


def test_update_calendar_event_rejects_blank_subject():
    with pytest.raises(OutlookSkillError, match="non-empty"):
        calendar.update_calendar_event(settings(), event_id="EVT_123", subject="   ")


def test_update_calendar_event_rejects_blank_location():
    with pytest.raises(OutlookSkillError, match="non-empty"):
        calendar.update_calendar_event(settings(), event_id="EVT_123", location="  ")


def test_update_calendar_event_rejects_negative_reminder():
    with pytest.raises(OutlookSkillError, match="zero or positive"):
        calendar.update_calendar_event(settings(), event_id="EVT_123", reminder_minutes=-5)


def test_update_calendar_event_reminder_zero_is_sent(monkeypatch):
    transport = FakeUpdateTransport()
    install_fake_calendar_graph(monkeypatch, transport)

    result = calendar.update_calendar_event(
        settings(),
        event_id="EVT_123",
        reminder_minutes=0,
    )

    assert result["changed_fields"] == ["--reminder-minutes"]
    graph_payload = json.loads(transport.requests[0][2])
    assert graph_payload == {"reminderMinutesBeforeStart": 0}


def test_update_calendar_event_location_only_payload(monkeypatch):
    transport = FakeUpdateTransport()
    install_fake_calendar_graph(monkeypatch, transport)

    result = calendar.update_calendar_event(settings(), event_id="EVT_123", location="Room B")

    assert result["changed_fields"] == ["--location"]
    graph_payload = json.loads(transport.requests[0][2])
    assert graph_payload == {"location": {"displayName": "Room B"}}


def test_update_calendar_event_rejects_nonexistent_spring_forward_time():
    with pytest.raises(OutlookSkillError, match="calendar update"):
        calendar.update_calendar_event(
            settings(),
            event_id="EVT_123",
            start="2026-03-08T02:30:00",
            end="2026-03-08T03:00:00",
            timezone="PT",
        )


def test_update_calendar_event_unknown_timezone_error_is_namespaced():
    with pytest.raises(OutlookSkillError, match="calendar update"):
        calendar.update_calendar_event(
            settings(),
            event_id="EVT_123",
            start="2026-04-15T13:00:00",
            end="2026-04-15T14:30:00",
            timezone="Mars/Olympus_Mons",
        )


def test_update_calendar_event_graph_error_raises(monkeypatch):
    class FailingTransport(httpx.BaseTransport):
        def handle_request(self, request: httpx.Request) -> httpx.Response:
            return httpx.Response(404, json={"error": {"message": "not found"}})

    install_fake_calendar_graph(monkeypatch, FailingTransport())

    with pytest.raises(calendar.GraphApiError):
        calendar.update_calendar_event(settings(), event_id="EVT_123", subject="Updated")
