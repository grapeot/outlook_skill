from __future__ import annotations

from email import policy
from email.message import EmailMessage, Message
from email.parser import BytesParser
from pathlib import Path
from typing import Any, Protocol, cast

from .config import Settings
from .errors import OutlookSkillError
from .graph_client import GraphMailClient
from .markdown_renderer import html_part_to_markdown, select_body_parts
from .store import MailStore


class ProgressCallback(Protocol):
    def __call__(self, *, current: int, total: int, downloaded: int, skipped: int) -> None: ...


def download_recent_mail(
    settings: Settings,
    *,
    days: int,
    limit: int,
    folders: tuple[str, ...] | None = None,
    progress_callback: ProgressCallback | None = None,
) -> dict[str, object]:
    target_folders = folders or settings.default_folders
    store = MailStore(settings)
    client = GraphMailClient(settings)
    try:
        downloaded: list[dict[str, object]] = []
        skipped = 0
        per_folder_messages: list[tuple[str, list[dict[str, str | None]]]] = []
        for folder_name in target_folders:
            folder_id, folder_display_name = client.resolve_folder(folder_name)
            messages = client.list_recent_messages(folder_id=folder_id, days=days, limit=limit)
            per_folder_messages.append((folder_display_name, messages))

        total = sum(len(messages) for _, messages in per_folder_messages)
        if progress_callback is not None:
            progress_callback(current=0, total=total, downloaded=0, skipped=0)
        current = 0
        for folder_display_name, messages in per_folder_messages:
            for message in messages:
                current += 1
                graph_id = message["id"]
                if not isinstance(graph_id, str):
                    if progress_callback is not None:
                        progress_callback(current=current, total=total, downloaded=len(downloaded), skipped=skipped)
                    continue
                if store.has_message(account=settings.email, folder=folder_display_name, uidvalidity="graph", uid=graph_id):
                    skipped += 1
                    if progress_callback is not None:
                        progress_callback(current=current, total=total, downloaded=len(downloaded), skipped=skipped)
                    continue
                raw_message = client.fetch_message_mime(message_id=graph_id)
                metadata = extract_metadata(raw_message)
                saved = store.save_message(
                    account=settings.email,
                    folder=folder_display_name,
                    uid=graph_id,
                    uidvalidity="graph",
                    raw_message=raw_message,
                    message_id=metadata["message_id"] or message.get("internet_message_id"),
                    subject=metadata["subject"],
                    from_addr=metadata["from_addr"],
                    received_at=metadata["received_at"] or message.get("received_at"),
                    internaldate=message.get("received_at"),
                    size=len(raw_message),
                    flags_json="",
                )
                downloaded.append(
                    {
                        "folder": folder_display_name,
                        "graph_id": saved.uid,
                        "subject": saved.subject,
                        "from_addr": saved.from_addr,
                        "received_at": saved.received_at,
                        "mime_path": saved.mime_path,
                    }
                )
                if progress_callback is not None:
                    progress_callback(current=current, total=total, downloaded=len(downloaded), skipped=skipped)
        return {
            "folders": [folder_name for folder_name, _ in per_folder_messages],
            "days": days,
            "downloaded_count": len(downloaded),
            "skipped_existing_count": skipped,
            "messages": downloaded,
        }
    finally:
        client.close()
        store.close()


def list_local_mail(settings: Settings, *, limit: int) -> dict[str, object]:
    store = MailStore(settings)
    try:
        messages = store.list_local_messages(limit=limit)
        return {"count": len(messages), "messages": messages}
    finally:
        store.close()


DEFAULT_BODY_CHAR_LIMIT = 10000


def read_local_mail(
    settings: Settings,
    *,
    subject: str | None = None,
    from_filter: str | None = None,
    graph_id: str | None = None,
    index: int | None = None,
    latest: bool = False,
    full: bool = False,
    body_char_limit: int = DEFAULT_BODY_CHAR_LIMIT,
) -> dict[str, object]:
    if not subject and not graph_id:
        raise OutlookSkillError("read requires --subject or --graph-id.")

    store = MailStore(settings)
    try:
        matches = store.find_messages(
            subject_substring=subject,
            from_substring=from_filter,
            graph_id=graph_id,
        )
    finally:
        store.close()

    if not matches:
        raise OutlookSkillError("No local messages matched the given filters.")

    if graph_id is not None:
        chosen = matches[0]
    elif len(matches) == 1:
        chosen = matches[0]
    elif latest:
        chosen = matches[0]
    elif index is not None:
        if index < 0 or index >= len(matches):
            raise OutlookSkillError(
                f"--index {index} out of range (matched {len(matches)} messages)."
            )
        chosen = matches[index]
    else:
        return {
            "match_count": len(matches),
            "matches": [
                {
                    "index": i,
                    "graph_id": m["uid"],
                    "subject": m["subject"],
                    "from_addr": m["from_addr"],
                    "received_at": m["received_at"],
                    "folder": m["folder"],
                }
                for i, m in enumerate(matches)
            ],
            "body": None,
            "note": "multiple matches; rerun with --latest, --index N, or --graph-id",
        }

    mime_path_value = chosen.get("mime_path")
    if not isinstance(mime_path_value, str):
        raise OutlookSkillError("Matched message has no mime_path; cannot read body.")
    mime_path = Path(mime_path_value)
    if not mime_path.exists():
        raise OutlookSkillError(f"MIME file missing on disk: {mime_path}")

    raw = mime_path.read_bytes()
    message = cast(EmailMessage, BytesParser(policy=policy.default).parsebytes(raw))
    body_text, body_source = extract_body_text(message)

    truncated = False
    truncated_remaining = 0
    if not full and len(body_text) > body_char_limit:
        truncated_remaining = len(body_text) - body_char_limit
        body_text = body_text[:body_char_limit]
        truncated = True

    return {
        "match_count": len(matches),
        "graph_id": chosen["uid"],
        "folder": chosen["folder"],
        "subject": header_string(message, "Subject") or chosen.get("subject"),
        "from": header_string(message, "From") or chosen.get("from_addr"),
        "to": header_string(message, "To"),
        "cc": header_string(message, "Cc"),
        "date": header_string(message, "Date") or chosen.get("received_at"),
        "mime_path": str(mime_path),
        "body_source": body_source,
        "body": body_text,
        "body_truncated": truncated,
        "body_truncated_remaining_chars": truncated_remaining,
    }


def extract_body_text(message: EmailMessage) -> tuple[str, str]:
    plain_part, html_part = select_body_parts(message)
    if plain_part is not None:
        text = _part_text(plain_part).strip()
        if text:
            return text, "plain"
    if html_part is not None:
        html_text = _part_text(html_part)
        markdown = html_part_to_markdown(html_text).strip()
        if markdown:
            return markdown, "html-converted"
    return "(no plain text or html body available)", "none"


def _part_text(part: EmailMessage | Message) -> str:
    try:
        content = cast(Any, part).get_content()
        return content if isinstance(content, str) else str(content)
    except Exception:
        payload = cast(bytes, part.get_payload(decode=True) or b"")
        charset = part.get_content_charset() or "utf-8"
        return payload.decode(charset, errors="replace")


def header_string(message: Message, key: str) -> str | None:
    value = message.get(key)
    if value is None:
        return None
    return str(value).strip() or None


def extract_metadata(raw_message: bytes) -> dict[str, str | None]:
    message = BytesParser(policy=policy.default).parsebytes(raw_message)
    return {
        "message_id": normalize_header(message.get("Message-ID")),
        "subject": normalize_header(message.get("Subject")),
        "from_addr": extract_from_address(message.get("From")),
        "received_at": normalize_header(message.get("Date")),
    }


def normalize_header(value: str | None) -> str | None:
    if value is None:
        return None
    stripped = value.strip()
    return stripped or None


def extract_from_address(value: str | None) -> str | None:
    normalized = normalize_header(value)
    if not normalized:
        return None
    if "<" in normalized and ">" in normalized:
        return normalized.split("<", 1)[1].split(">", 1)[0].strip()
    return normalized
