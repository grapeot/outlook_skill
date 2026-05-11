from __future__ import annotations

import json
from typing import Any, cast
from urllib.parse import unquote

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
    assert transport.requests[0][0] == "POST"
    assert transport.requests[0][1].endswith("/me/calendar/events")
    graph_payload = json.loads(transport.requests[0][2])
    assert graph_payload["subject"] == "Meeting"
    assert graph_payload["start"] == {"dateTime": "2026-05-06T10:00:00", "timeZone": "Pacific Standard Time"}
    assert graph_payload["end"] == {"dateTime": "2026-05-06T10:30:00", "timeZone": "Pacific Standard Time"}
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


def test_create_calendar_invite_rejects_missing_attendees():
    with pytest.raises(OutlookSkillError):
        calendar.create_calendar_invite(
            settings(),
            subject="Meeting",
            start="2026-05-06T10:00:00",
            end="2026-05-06T10:30:00",
            attendees=(),
        )


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
