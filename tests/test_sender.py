from __future__ import annotations

import json
from typing import Any, cast

import httpx
import pytest

from outlook_skill import sender
from outlook_skill.config import Settings
from outlook_skill.errors import OutlookSkillError


class FakeTransport(httpx.BaseTransport):
    def __init__(self) -> None:
        self.requests: list[tuple[str, str, Any]] = []

    def handle_request(self, request: httpx.Request) -> httpx.Response:
        self.requests.append((request.method, str(request.url), request.content))
        if request.method == "POST" and str(request.url).endswith("/me/sendMail"):
            return httpx.Response(202)
        if request.method == "POST" and str(request.url).endswith("/me/messages"):
            payload = json.loads(request.content)
            return httpx.Response(201, json={
                "id": "DRAFT_123",
                "subject": payload.get("subject"),
                "webLink": "https://outlook.live.com/draft/DRAFT_123",
                "parentFolderId": "DRAFTS_FOLDER",
            })
        return httpx.Response(404, json={"error": "unexpected"})


def install_fake_graph(monkeypatch, transport: FakeTransport) -> None:
    class FakeAuthManager:
        def __init__(self, settings): pass
        def get_access_token(self): return "TOKEN"

    original_client = httpx.Client

    def fake_client(*args, **kwargs):
        kwargs.pop("transport", None)
        return original_client(*args, transport=transport, **kwargs)

    monkeypatch.setattr(sender, "AuthManager", FakeAuthManager)
    monkeypatch.setattr(sender.httpx, "Client", fake_client)


def settings() -> Settings:
    value = object.__new__(type("S", (), {"graph_base_url": "https://example.test/v1.0"}))
    return cast(Settings, cast(object, value))


def test_send_mail_dry_run_does_not_call_graph(monkeypatch):
    transport = FakeTransport()
    install_fake_graph(monkeypatch, transport)

    payload = sender.send_mail(
        settings(),
        to=("duck@example.com",),
        subject="Hello",
        body_text="Body",
        dry_run=True,
    )

    assert payload["dry_run"] is True
    assert payload["sent"] is False
    assert transport.requests == []


def test_send_mail_posts_to_me_sendmail(monkeypatch):
    transport = FakeTransport()
    install_fake_graph(monkeypatch, transport)

    payload = sender.send_mail(
        settings(),
        to=("duck@example.com",),
        cc=("cc@example.com",),
        bcc=("bcc@example.com",),
        subject="Hello",
        body_text="Body",
    )

    assert payload["sent"] is True
    assert transport.requests[0][0] == "POST"
    assert transport.requests[0][1].endswith("/me/sendMail")
    graph_payload = json.loads(transport.requests[0][2])
    message = graph_payload["message"]
    assert message["subject"] == "Hello"
    assert message["body"] == {"contentType": "Text", "content": "Body"}
    assert message["toRecipients"][0]["emailAddress"]["address"] == "duck@example.com"
    assert message["ccRecipients"][0]["emailAddress"]["address"] == "cc@example.com"
    assert message["bccRecipients"][0]["emailAddress"]["address"] == "bcc@example.com"


def test_send_mail_markdown_body_converts_to_html(monkeypatch):
    transport = FakeTransport()
    install_fake_graph(monkeypatch, transport)

    sender.send_mail(
        settings(),
        to=("duck@example.com",),
        subject="Hello",
        body_text="**bold**",
        body_format="markdown",
    )

    graph_payload = json.loads(transport.requests[0][2])
    body = graph_payload["message"]["body"]
    assert body["contentType"] == "HTML"
    assert "<strong>bold</strong>" in body["content"]


def test_send_mail_includes_small_attachment(monkeypatch, tmp_path):
    transport = FakeTransport()
    install_fake_graph(monkeypatch, transport)
    attachment = tmp_path / "note.txt"
    attachment.write_text("hello", encoding="utf-8")

    sender.send_mail(
        settings(),
        to=("duck@example.com",),
        subject="Hello",
        body_text="Body",
        attachments=(attachment,),
    )

    graph_payload = json.loads(transport.requests[0][2])
    attachment_payload = graph_payload["message"]["attachments"][0]
    assert attachment_payload["name"] == "note.txt"
    assert attachment_payload["contentBytes"] == "aGVsbG8="


def test_send_mail_rejects_missing_to():
    with pytest.raises(OutlookSkillError):
        sender.send_mail(settings(), to=(), subject="Hello", body_text="Body")


def test_send_mail_rejects_missing_subject():
    with pytest.raises(OutlookSkillError):
        sender.send_mail(settings(), to=("duck@example.com",), subject=" ", body_text="Body")


def test_send_mail_rejects_bad_body_format():
    with pytest.raises(OutlookSkillError):
        sender.send_mail(settings(), to=("duck@example.com",), subject="Hello", body_text="Body", body_format="xml")


def test_send_mail_rejects_missing_attachment(tmp_path):
    with pytest.raises(OutlookSkillError):
        sender.send_mail(
            settings(),
            to=("duck@example.com",),
            subject="Hello",
            body_text="Body",
            attachments=(tmp_path / "missing.txt",),
        )


def test_send_mail_rejects_large_attachment(monkeypatch, tmp_path):
    monkeypatch.setattr(sender, "SMALL_ATTACHMENT_LIMIT", 4)
    attachment = tmp_path / "large.txt"
    attachment.write_text("hello", encoding="utf-8")

    with pytest.raises(OutlookSkillError):
        sender.send_mail(
            settings(),
            to=("duck@example.com",),
            subject="Hello",
            body_text="Body",
            attachments=(attachment,),
        )


def test_create_mail_draft_posts_to_me_messages_without_recipients(monkeypatch):
    transport = FakeTransport()
    install_fake_graph(monkeypatch, transport)

    payload = sender.create_mail_draft(
        settings(),
        subject="Draft subject",
        body_text="Draft body",
    )

    assert payload["operation"] == "draft"
    assert payload["created"] is True
    assert payload["sent"] is False
    assert payload["draft_id"] == "DRAFT_123"
    assert payload["to"] == []
    assert payload["cc"] == []
    assert transport.requests[0][0] == "POST"
    assert transport.requests[0][1].endswith("/me/messages")
    graph_payload = json.loads(transport.requests[0][2])
    assert graph_payload["subject"] == "Draft subject"
    assert graph_payload["body"] == {"contentType": "Text", "content": "Draft body"}
    assert graph_payload["toRecipients"] == []
    assert graph_payload["ccRecipients"] == []
    assert graph_payload["bccRecipients"] == []


def test_create_mail_draft_accepts_recipients_and_markdown(monkeypatch):
    transport = FakeTransport()
    install_fake_graph(monkeypatch, transport)

    sender.create_mail_draft(
        settings(),
        to=("duck@example.com",),
        cc=("cc@example.com",),
        subject="Draft subject",
        body_text="**bold**",
        body_format="markdown",
    )

    graph_payload = json.loads(transport.requests[0][2])
    assert graph_payload["toRecipients"][0]["emailAddress"]["address"] == "duck@example.com"
    assert graph_payload["ccRecipients"][0]["emailAddress"]["address"] == "cc@example.com"
    assert graph_payload["body"]["contentType"] == "HTML"
    assert "<strong>bold</strong>" in graph_payload["body"]["content"]


def test_create_mail_draft_rejects_missing_subject():
    with pytest.raises(OutlookSkillError):
        sender.create_mail_draft(settings(), subject=" ", body_text="Body")
