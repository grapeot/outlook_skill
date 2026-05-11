from __future__ import annotations

import base64
import mimetypes
from pathlib import Path

import httpx
from markdown import markdown as md_to_html

from .auth import AuthManager
from .config import Settings
from .errors import GraphApiError, OutlookSkillError
from .replier import SMALL_ATTACHMENT_LIMIT


def send_mail(
    settings: Settings,
    *,
    to: tuple[str, ...],
    subject: str,
    body_text: str,
    body_format: str = "text",
    cc: tuple[str, ...] = (),
    bcc: tuple[str, ...] = (),
    attachments: tuple[Path, ...] = (),
    save_to_sent_items: bool = True,
    dry_run: bool = False,
) -> dict[str, object]:
    if not to:
        raise OutlookSkillError("send requires at least one --to recipient.")
    if not subject.strip():
        raise OutlookSkillError("send requires --subject.")
    if body_format not in ("text", "html", "markdown", "md"):
        raise OutlookSkillError(f"Unsupported body format: {body_format}")

    for path in attachments:
        if not path.exists():
            raise OutlookSkillError(f"Attachment not found: {path}")
        if path.stat().st_size > SMALL_ATTACHMENT_LIMIT:
            raise OutlookSkillError("mail send currently supports attachments up to 3 MB; use reply for large-attachment draft upload support.")

    content_type, content_value = _prepare_body(body_text, body_format)
    attachment_payloads = [_small_attachment_payload(path) for path in attachments]
    message: dict[str, object] = {
        "subject": subject,
        "body": {"contentType": content_type, "content": content_value},
        "toRecipients": [_recipient(addr) for addr in to],
    }
    if cc:
        message["ccRecipients"] = [_recipient(addr) for addr in cc]
    if bcc:
        message["bccRecipients"] = [_recipient(addr) for addr in bcc]
    if attachment_payloads:
        message["attachments"] = attachment_payloads
    graph_payload: dict[str, object] = {
        "message": message,
        "saveToSentItems": save_to_sent_items,
    }

    result = {
        "dry_run": dry_run,
        "endpoint": "/me/sendMail",
        "subject": subject,
        "to": list(to),
        "cc": list(cc),
        "bcc": list(bcc),
        "attachment_count": len(attachment_payloads),
        "attachments": [
            {"name": path.name, "size": path.stat().st_size, "content_type": mimetypes.guess_type(path.name)[0] or "application/octet-stream"}
            for path in attachments
        ],
        "body_content_type": content_type,
        "body_chars": len(content_value),
        "save_to_sent_items": save_to_sent_items,
    }
    if dry_run:
        return {**result, "sent": False, "note": "dry run; Graph /me/sendMail was not called"}

    token = AuthManager(settings).get_access_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Content-Type": "application/json",
    }
    with httpx.Client(base_url=settings.graph_base_url, timeout=httpx.Timeout(120.0, connect=10.0), headers=headers) as client:
        response = client.post("/me/sendMail", json=graph_payload)
    if response.is_error:
        raise GraphApiError(
            f"Graph sendMail failed with status {response.status_code}",
            status_code=response.status_code,
            response_text=response.text,
        )
    return {**result, "sent": True}


def _prepare_body(body_text: str, body_format: str) -> tuple[str, str]:
    if body_format in ("markdown", "md"):
        return "HTML", md_to_html(body_text, extensions=["extra", "sane_lists"])
    if body_format == "html":
        return "HTML", body_text
    return "Text", body_text


def _recipient(address: str) -> dict[str, dict[str, str]]:
    return {"emailAddress": {"address": address}}


def _small_attachment_payload(path: Path) -> dict[str, object]:
    content_type, _ = mimetypes.guess_type(path.name)
    content_type = content_type or "application/octet-stream"
    return {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": path.name,
        "contentType": content_type,
        "contentBytes": base64.b64encode(path.read_bytes()).decode("ascii"),
    }
