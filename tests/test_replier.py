from __future__ import annotations

from typing import Any, cast

import httpx
import pytest

from outlook_skill import replier
from outlook_skill.config import Settings
from outlook_skill.errors import OutlookSkillError


class FakeTransport(httpx.BaseTransport):
    def __init__(self, create_action: str = "createReply") -> None:
        self.requests: list[tuple[str, str, Any]] = []
        self.draft_id: str = "DRAFT_123"
        self.create_action: str = create_action

    def handle_request(self, request: httpx.Request) -> httpx.Response:
        self.requests.append((request.method, str(request.url), request.content))
        url = str(request.url)
        if request.method == "POST" and url.endswith(f"/{self.create_action}"):
            return httpx.Response(201, json={"id": self.draft_id, "subject": "Re: Test"})
        if request.method == "PATCH" and f"/me/messages/{self.draft_id}" in url:
            return httpx.Response(200, json={"id": self.draft_id})
        if request.method == "POST" and url.endswith(f"/me/messages/{self.draft_id}/attachments"):
            return httpx.Response(201, json={"id": "ATTACH_1"})
        if request.method == "GET" and url.endswith(f"/me/messages/{self.draft_id}"):
            return httpx.Response(
                200,
                json={
                    "subject": "Re: Test",
                    "toRecipients": [{"emailAddress": {"address": "reply@example.com"}}],
                    "ccRecipients": [],
                },
            )
        if request.method == "POST" and url.endswith(f"/me/messages/{self.draft_id}/send"):
            return httpx.Response(202)
        return httpx.Response(404, json={"error": "unexpected"})


def install_fake_graph(monkeypatch, transport: FakeTransport) -> None:
    class FakeAuthManager:
        def __init__(self, settings): pass
        def get_access_token(self): return "TOKEN"

    original_client = httpx.Client

    def fake_client(*args, **kwargs):
        kwargs.pop("transport", None)
        return original_client(*args, transport=transport, **kwargs)

    monkeypatch.setattr(replier, "AuthManager", FakeAuthManager)
    monkeypatch.setattr(replier.httpx, "Client", fake_client)


def settings() -> Settings:
    value = object.__new__(type("S", (), {"graph_base_url": "https://example.test/v1.0"}))
    return cast(Settings, cast(object, value))


def test_reply_dry_run_creates_draft_with_attachment(monkeypatch, tmp_path):
    transport = FakeTransport()
    install_fake_graph(monkeypatch, transport)

    attach = tmp_path / "note.txt"
    attach.write_text("hello")

    payload = replier.reply_to_message(
        settings(),
        graph_id="ORIG_ID",
        body_text="Hi there",
        body_format="text",
        attachments=(attach,),
        dry_run=True,
    )

    assert payload["dry_run"] is True
    assert payload["operation"] == "reply"
    assert payload["draft_id"] == transport.draft_id
    assert payload["attachment_count"] == 1
    methods = [m for m, _, _ in transport.requests]
    assert "POST" in methods  # createReply
    assert "PATCH" in methods  # body update
    # Dry run: no /send call
    assert not any(url.endswith("/send") for _, url, _ in transport.requests)


def test_reply_all_dry_run_uses_create_reply_all_and_does_not_send(monkeypatch):
    transport = FakeTransport(create_action="createReplyAll")
    install_fake_graph(monkeypatch, transport)

    payload = replier.reply_to_message(
        settings(),
        graph_id="ORIG_ID",
        body_text="Hi all",
        dry_run=True,
        reply_all=True,
    )

    assert payload["dry_run"] is True
    assert payload["operation"] == "reply_all"
    assert any(url.endswith("/createReplyAll") for _, url, _ in transport.requests)
    assert not any(url.endswith("/createReply") for _, url, _ in transport.requests)
    assert not any(url.endswith("/send") for _, url, _ in transport.requests)


def test_reply_rejects_missing_graph_id():
    with pytest.raises(OutlookSkillError):
        replier.reply_to_message(settings(), graph_id="", body_text="hi")


def test_reply_rejects_bad_format():
    with pytest.raises(OutlookSkillError):
        replier.reply_to_message(settings(), graph_id="X", body_text="hi", body_format="xml")


def test_reply_rejects_missing_attachment(tmp_path):
    with pytest.raises(OutlookSkillError):
        replier.reply_to_message(
            settings(),
            graph_id="X",
            body_text="hi",
            attachments=(tmp_path / "nope.txt",),
        )


def test_prepare_body_markdown_to_html():
    content_type, content = replier._prepare_body("**bold**", "markdown")
    assert content_type == "HTML"
    assert "<strong>bold</strong>" in content


def test_prepare_body_text_passthrough():
    content_type, content = replier._prepare_body("plain", "text")
    assert content_type == "Text"
    assert content == "plain"
