from __future__ import annotations

import base64
import mimetypes
from pathlib import Path
from typing import Any, cast

import httpx
from markdown import markdown as md_to_html

from .auth import AuthManager
from .config import Settings
from .errors import GraphApiError, OutlookSkillError


SMALL_ATTACHMENT_LIMIT = 3 * 1024 * 1024  # 3 MB: above this, Graph requires upload sessions.


def reply_to_message(
    settings: Settings,
    *,
    graph_id: str,
    body_text: str,
    body_format: str = "text",
    attachments: tuple[Path, ...] = (),
    to_override: tuple[str, ...] = (),
    cc_override: tuple[str, ...] = (),
    dry_run: bool = False,
) -> dict[str, object]:
    if not graph_id:
        raise OutlookSkillError("reply requires --graph-id.")
    if body_format not in ("text", "html", "markdown", "md"):
        raise OutlookSkillError(f"Unsupported body format: {body_format}")

    for path in attachments:
        if not path.exists():
            raise OutlookSkillError(f"Attachment not found: {path}")

    content_type, content_value = _prepare_body(body_text, body_format)

    token = AuthManager(settings).get_access_token()
    base_url = settings.graph_base_url
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    }
    with httpx.Client(base_url=base_url, timeout=httpx.Timeout(120.0, connect=10.0), headers=headers) as client:
        draft = _create_reply_draft(client, graph_id=graph_id)
        draft_id = cast(str, draft["id"])

        patch_payload: dict[str, Any] = {
            "body": {"contentType": content_type, "content": content_value},
        }
        if to_override:
            patch_payload["toRecipients"] = [_recipient(addr) for addr in to_override]
        if cc_override:
            patch_payload["ccRecipients"] = [_recipient(addr) for addr in cc_override]

        _patch_json(client, f"/me/messages/{draft_id}", patch_payload)

        attachment_results: list[dict[str, object]] = []
        for path in attachments:
            info = _attach_file(client, draft_id=draft_id, file_path=path)
            attachment_results.append(info)

        draft_meta = _get_json(client, f"/me/messages/{draft_id}")
        to_recipients = _render_recipients(draft_meta.get("toRecipients"))
        cc_recipients = _render_recipients(draft_meta.get("ccRecipients"))
        subject = draft_meta.get("subject")

        if dry_run:
            return {
                "dry_run": True,
                "draft_id": draft_id,
                "subject": subject if isinstance(subject, str) else None,
                "to": to_recipients,
                "cc": cc_recipients,
                "attachment_count": len(attachment_results),
                "attachments": attachment_results,
                "body_content_type": content_type,
                "body_chars": len(content_value),
                "note": "draft created but not sent",
            }

        _post_empty(client, f"/me/messages/{draft_id}/send")

        return {
            "dry_run": False,
            "draft_id": draft_id,
            "subject": subject if isinstance(subject, str) else None,
            "to": to_recipients,
            "cc": cc_recipients,
            "attachment_count": len(attachment_results),
            "attachments": attachment_results,
            "body_content_type": content_type,
            "body_chars": len(content_value),
            "sent": True,
        }


def _prepare_body(body_text: str, body_format: str) -> tuple[str, str]:
    if body_format in ("markdown", "md"):
        html = md_to_html(body_text, extensions=["extra", "sane_lists"])
        return "HTML", html
    if body_format == "html":
        return "HTML", body_text
    return "Text", body_text


def _create_reply_draft(client: httpx.Client, *, graph_id: str) -> dict[str, Any]:
    response = client.post(f"/me/messages/{graph_id}/createReply", json={})
    if response.is_error:
        raise GraphApiError(
            f"Graph createReply failed with status {response.status_code}",
            status_code=response.status_code,
            response_text=response.text,
        )
    payload = response.json()
    if not isinstance(payload, dict) or "id" not in payload:
        raise GraphApiError("createReply returned an unexpected payload shape.")
    return cast(dict[str, Any], payload)


def _patch_json(client: httpx.Client, path: str, body: dict[str, Any]) -> dict[str, Any]:
    response = client.patch(path, json=body)
    if response.is_error:
        raise GraphApiError(
            f"Graph PATCH {path} failed with status {response.status_code}",
            status_code=response.status_code,
            response_text=response.text,
        )
    if not response.content:
        return {}
    payload = response.json()
    if not isinstance(payload, dict):
        raise GraphApiError("Graph PATCH returned an unexpected payload shape.")
    return cast(dict[str, Any], payload)


def _get_json(client: httpx.Client, path: str) -> dict[str, Any]:
    response = client.get(path)
    if response.is_error:
        raise GraphApiError(
            f"Graph GET {path} failed with status {response.status_code}",
            status_code=response.status_code,
            response_text=response.text,
        )
    payload = response.json()
    if not isinstance(payload, dict):
        raise GraphApiError("Graph GET returned an unexpected payload shape.")
    return cast(dict[str, Any], payload)


def _post_empty(client: httpx.Client, path: str) -> None:
    response = client.post(path)
    if response.is_error:
        raise GraphApiError(
            f"Graph POST {path} failed with status {response.status_code}",
            status_code=response.status_code,
            response_text=response.text,
        )


def _recipient(address: str) -> dict[str, Any]:
    return {"emailAddress": {"address": address}}


def _render_recipients(raw: Any) -> list[str]:
    if not isinstance(raw, list):
        return []
    out: list[str] = []
    for item in raw:
        if not isinstance(item, dict):
            continue
        email = item.get("emailAddress")
        if isinstance(email, dict):
            addr = email.get("address")
            if isinstance(addr, str):
                out.append(addr)
    return out


def _attach_file(client: httpx.Client, *, draft_id: str, file_path: Path) -> dict[str, object]:
    size = file_path.stat().st_size
    content_type, _ = mimetypes.guess_type(file_path.name)
    content_type = content_type or "application/octet-stream"

    if size <= SMALL_ATTACHMENT_LIMIT:
        content_bytes = file_path.read_bytes()
        payload = {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": file_path.name,
            "contentType": content_type,
            "contentBytes": base64.b64encode(content_bytes).decode("ascii"),
        }
        response = client.post(f"/me/messages/{draft_id}/attachments", json=payload)
        if response.is_error:
            raise GraphApiError(
                f"Graph attachment upload failed with status {response.status_code}",
                status_code=response.status_code,
                response_text=response.text,
            )
        result = response.json()
        attach_id = result.get("id") if isinstance(result, dict) else None
        return {
            "name": file_path.name,
            "size": size,
            "content_type": content_type,
            "mode": "inline",
            "id": attach_id if isinstance(attach_id, str) else None,
        }

    # Large attachment: upload session.
    session_payload = {
        "AttachmentItem": {
            "attachmentType": "file",
            "name": file_path.name,
            "size": size,
            "contentType": content_type,
        }
    }
    session_response = client.post(
        f"/me/messages/{draft_id}/attachments/createUploadSession",
        json=session_payload,
    )
    if session_response.is_error:
        raise GraphApiError(
            f"createUploadSession failed with status {session_response.status_code}",
            status_code=session_response.status_code,
            response_text=session_response.text,
        )
    session_data = session_response.json()
    upload_url = session_data.get("uploadUrl") if isinstance(session_data, dict) else None
    if not isinstance(upload_url, str):
        raise GraphApiError("Upload session missing uploadUrl.")

    chunk_size = 4 * 1024 * 1024  # 4 MB chunks.
    with file_path.open("rb") as handle, httpx.Client(timeout=httpx.Timeout(300.0, connect=10.0)) as uploader:
        offset = 0
        while offset < size:
            chunk = handle.read(chunk_size)
            if not chunk:
                break
            end = offset + len(chunk) - 1
            upload_headers = {
                "Content-Length": str(len(chunk)),
                "Content-Range": f"bytes {offset}-{end}/{size}",
            }
            put = uploader.put(upload_url, content=chunk, headers=upload_headers)
            if put.is_error:
                raise GraphApiError(
                    f"Upload session PUT failed with status {put.status_code}",
                    status_code=put.status_code,
                    response_text=put.text,
                )
            offset = end + 1

    return {
        "name": file_path.name,
        "size": size,
        "content_type": content_type,
        "mode": "upload-session",
        "id": None,
    }
