from __future__ import annotations

import re
from datetime import UTC, datetime, timedelta
from email import policy
from email.message import EmailMessage
from email.parser import BytesParser
from pathlib import Path
from typing import cast

from .config import Settings
from .downloader import extract_body_text, header_string
from .store import MailStore


SLUG_STRIP = re.compile(r"[^a-zA-Z0-9\u4e00-\u9fff\s-]+")
SLUG_SPACES = re.compile(r"\s+")


def export_local_to_markdown(
    settings: Settings,
    *,
    days: int | None = None,
    folders: tuple[str, ...] | None = None,
    subject: str | None = None,
    from_filter: str | None = None,
    force: bool = False,
    output_dir: Path | None = None,
) -> dict[str, object]:
    target_dir = output_dir or (settings.data_dir / "markdown")
    target_dir.mkdir(parents=True, exist_ok=True)

    store = MailStore(settings)
    try:
        matches = store.find_messages(
            subject_substring=subject,
            from_substring=from_filter,
        )
    finally:
        store.close()

    if folders:
        folder_set = {f.strip() for f in folders}
        matches = [m for m in matches if str(m.get("folder") or "") in folder_set]

    cutoff: datetime | None = None
    if days is not None:
        cutoff = datetime.now(UTC) - timedelta(days=days)

    exported: list[dict[str, object]] = []
    skipped_existing = 0
    skipped_filter = 0
    failed: list[dict[str, object]] = []

    for match in matches:
        received_at_raw = match.get("received_at") or match.get("internaldate") or match.get("downloaded_at")
        received_at_str = received_at_raw if isinstance(received_at_raw, str) else None
        received_dt = parse_datetime(received_at_str)
        if cutoff is not None and received_dt is not None and received_dt < cutoff:
            skipped_filter += 1
            continue

        mime_path_value = match.get("mime_path")
        if not isinstance(mime_path_value, str):
            skipped_filter += 1
            continue
        mime_path = Path(mime_path_value)
        if not mime_path.exists():
            failed.append({"graph_id": match.get("uid"), "error": "mime file missing"})
            continue

        graph_id = match.get("uid")
        graph_id_str = graph_id if isinstance(graph_id, str) else ""
        folder_name = str(match.get("folder") or "unknown")
        subject_raw = match.get("subject")
        subject_str = subject_raw if isinstance(subject_raw, str) else ""

        output_path = target_dir / build_md_filename(
            received_at=received_dt,
            folder=folder_name,
            subject=subject_str,
            graph_id=graph_id_str,
        )
        if output_path.exists() and not force:
            skipped_existing += 1
            continue

        try:
            raw = mime_path.read_bytes()
            message = cast(EmailMessage, BytesParser(policy=policy.default).parsebytes(raw))
            body_text, body_source = extract_body_text(message)
            md_content = render_markdown_document(
                graph_id=graph_id_str,
                folder=folder_name,
                message=message,
                db_subject=subject_str,
                db_from=str(match.get("from_addr") or "") or None,
                db_received_at=received_at_str,
                body=body_text,
                body_source=body_source,
            )
            output_path.write_text(md_content, encoding="utf-8")
            exported.append(
                {
                    "graph_id": graph_id_str,
                    "folder": folder_name,
                    "subject": subject_str,
                    "output_path": str(output_path),
                    "body_source": body_source,
                }
            )
        except Exception as exc:
            failed.append({"graph_id": graph_id_str, "error": str(exc)})

    return {
        "output_dir": str(target_dir),
        "exported_count": len(exported),
        "skipped_existing_count": skipped_existing,
        "skipped_filter_count": skipped_filter,
        "failed_count": len(failed),
        "failed": failed,
        "files": exported,
    }


def render_markdown_document(
    *,
    graph_id: str,
    folder: str,
    message: EmailMessage,
    db_subject: str | None,
    db_from: str | None,
    db_received_at: str | None,
    body: str,
    body_source: str,
) -> str:
    subject = header_string(message, "Subject") or db_subject or ""
    from_value = header_string(message, "From") or db_from or ""
    to_value = header_string(message, "To") or ""
    cc_value = header_string(message, "Cc") or ""
    date_value = header_string(message, "Date") or db_received_at or ""

    frontmatter_lines = [
        "---",
        f"graph_id: {yaml_escape(graph_id)}",
        f"folder: {yaml_escape(folder)}",
        f"from: {yaml_escape(from_value)}",
        f"to: {yaml_escape(to_value)}",
        f"cc: {yaml_escape(cc_value)}",
        f"subject: {yaml_escape(subject)}",
        f"date: {yaml_escape(date_value)}",
        f"body_source: {yaml_escape(body_source)}",
        "---",
        "",
    ]
    return "\n".join(frontmatter_lines) + body.rstrip() + "\n"


def yaml_escape(value: str) -> str:
    if value is None:
        return '""'
    # Quote values that contain special YAML characters.
    needs_quote = any(ch in value for ch in (":", "#", "\n", '"', "'", "\\"))
    if not needs_quote:
        return value
    escaped = value.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n")
    return f'"{escaped}"'


def parse_datetime(value: str | None) -> datetime | None:
    if not value:
        return None
    text = value.strip()
    if not text:
        return None
    # Try ISO 8601 first.
    try:
        normalized = text.replace("Z", "+00:00")
        parsed = datetime.fromisoformat(normalized)
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=UTC)
        return parsed.astimezone(UTC)
    except ValueError:
        pass
    # Fall back to RFC-2822 style.
    try:
        from email.utils import parsedate_to_datetime

        parsed = parsedate_to_datetime(text)
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=UTC)
        return parsed.astimezone(UTC)
    except (TypeError, ValueError):
        return None


def build_md_filename(
    *,
    received_at: datetime | None,
    folder: str,
    subject: str,
    graph_id: str,
) -> str:
    date_part = received_at.strftime("%Y-%m-%d") if received_at else "0000-00-00"
    folder_part = slugify(folder, limit=20) or "folder"
    subject_part = slugify(subject, limit=50) or "no-subject"
    short_id = short_graph_id(graph_id)
    return f"{date_part}_{folder_part}_{subject_part}_{short_id}.md"


def slugify(text: str, *, limit: int) -> str:
    if not text:
        return ""
    cleaned = SLUG_STRIP.sub(" ", text)
    cleaned = SLUG_SPACES.sub("-", cleaned.strip())
    cleaned = cleaned.strip("-")
    if len(cleaned) > limit:
        cleaned = cleaned[:limit].rstrip("-")
    return cleaned


def short_graph_id(graph_id: str) -> str:
    if not graph_id:
        return "00000000"
    cleaned = re.sub(r"[^A-Za-z0-9]", "", graph_id)
    if len(cleaned) < 8:
        return (cleaned + "0" * 8)[:8]
    return cleaned[-8:]


