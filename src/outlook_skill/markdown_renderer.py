from __future__ import annotations

import concurrent.futures
import re
import urllib.parse
from dataclasses import dataclass
from email import policy
from email.message import EmailMessage, Message
from email.parser import BytesParser
from pathlib import Path
from typing import Any, Protocol, cast

from bs4 import BeautifulSoup
from markdownify import markdownify as html_to_markdown

from .config import Settings


QUOTE_SEPARATORS = (
    re.compile(r"^On .+wrote:$", re.IGNORECASE),
    re.compile(r"^From:\s", re.IGNORECASE),
    re.compile(r"^Sent:\s", re.IGNORECASE),
    re.compile(r"^To:\s", re.IGNORECASE),
    re.compile(r"^Subject:\s", re.IGNORECASE),
    re.compile(r"^-{2,}\s*Original Message\s*-{2,}$", re.IGNORECASE),
)
ZERO_WIDTH_CHARS = re.compile(r"[\u200b\u200c\u200d\ufeff\u2060]+")
TRACKING_HOST_MARKERS = (
    "mailgun",
    "bloomreach",
    "klaviyo",
    "mandrillapp",
    "list-manage.com",
    "sendgrid.net",
    "emltrk.com",
)


class RenderProgressCallback(Protocol):
    def __call__(self, *, current: int, total: int, rendered: int, skipped: int) -> None: ...


@dataclass(frozen=True)
class AttachmentInfo:
    filename: str | None
    content_type: str
    size: int
    disposition: str | None
    content_id: str | None


@dataclass(frozen=True)
class RenderedMarkdown:
    source_path: Path
    output_path: Path
    subject: str | None
    attachment_count: int


def render_maildir_to_markdown(
    settings: Settings,
    *,
    input_dir: Path | None = None,
    output_dir: Path | None = None,
    limit: int | None = None,
    sample_size: int | None = None,
    seed: int = 0,
    workers: int = 8,
    progress_callback: RenderProgressCallback | None = None,
) -> dict[str, object]:
    source_dir = input_dir or settings.messages_dir
    target_dir = output_dir or settings.markdowns_dir
    source_dir.mkdir(parents=True, exist_ok=True)
    target_dir.mkdir(parents=True, exist_ok=True)

    eml_files = sorted(source_dir.glob("*.eml"))
    if limit is not None:
        eml_files = eml_files[:limit]
    if sample_size is not None and sample_size < len(eml_files):
        import random

        rng = random.Random(seed)
        eml_files = rng.sample(eml_files, sample_size)

    total = len(eml_files)
    rendered_count = 0
    skipped = 0
    outputs: list[dict[str, object]] = []

    if progress_callback is not None:
        progress_callback(current=0, total=total, rendered=0, skipped=0)

    with concurrent.futures.ThreadPoolExecutor(max_workers=max(1, workers)) as executor:
        future_map = {
            executor.submit(render_eml_file, path, target_dir): path
            for path in eml_files
        }
        current = 0
        for future in concurrent.futures.as_completed(future_map):
            current += 1
            try:
                result = future.result()
            except Exception:
                result = None
            if result is None:
                skipped += 1
            else:
                rendered_count += 1
                outputs.append(
                    {
                        "source_path": str(result.source_path),
                        "output_path": str(result.output_path),
                        "subject": result.subject,
                        "attachment_count": result.attachment_count,
                    }
                )
            if progress_callback is not None:
                progress_callback(current=current, total=total, rendered=rendered_count, skipped=skipped)

    outputs.sort(key=lambda item: cast(str, item["output_path"]))
    return {
        "input_dir": str(source_dir),
        "output_dir": str(target_dir),
        "total_files": total,
        "rendered_count": rendered_count,
        "skipped_count": skipped,
        "files": outputs,
    }


def render_eml_file(source_path: Path, output_dir: Path) -> RenderedMarkdown | None:
    raw = source_path.read_bytes()
    message = cast(EmailMessage, BytesParser(policy=policy.default).parsebytes(raw))
    markdown = render_message_to_markdown(message, source_name=source_path.name)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{source_path.stem}.md"
    output_path.write_text(cast(str, markdown["markdown"]), encoding="utf-8")
    return RenderedMarkdown(
        source_path=source_path,
        output_path=output_path,
        subject=cast(str | None, markdown["subject"]),
        attachment_count=len(cast(list[AttachmentInfo], markdown["attachments"])),
    )


def render_message_to_markdown(message: EmailMessage, *, source_name: str) -> dict[str, object]:
    subject = header_value(message, "Subject")
    from_value = header_value(message, "From")
    to_value = header_value(message, "To")
    cc_value = header_value(message, "Cc")
    date_value = header_value(message, "Date")
    message_id = header_value(message, "Message-ID")
    body_markdown = render_body_markdown(message)
    attachments = collect_attachments(message)

    lines = [
        f"# {subject or '(no subject)'}",
        "",
        f"- Source: `{source_name}`",
        f"- From: {from_value or ''}",
        f"- To: {to_value or ''}",
        f"- Cc: {cc_value or ''}",
        f"- Date: {date_value or ''}",
        f"- Message-ID: {message_id or ''}",
        f"- Attachment count: {len(attachments)}",
        "",
        "## Body",
        "",
        body_markdown or "(empty body)",
    ]

    if attachments:
        lines.extend(["", "## Attachments", ""])
        for attachment in attachments:
            name = attachment.filename or "(unnamed)"
            cid = f", cid={attachment.content_id}" if attachment.content_id else ""
            disposition = attachment.disposition or "unknown"
            lines.append(
                f"- `{name}` | type={attachment.content_type} | size={attachment.size} | disposition={disposition}{cid}"
            )

    markdown = "\n".join(lines).strip() + "\n"
    return {
        "subject": subject,
        "markdown": markdown,
        "attachments": attachments,
    }


def render_body_markdown(message: EmailMessage) -> str:
    plain_part, html_part = select_body_parts(message)

    plain_text = part_text(plain_part)
    html_text = html_part_to_markdown(part_text(html_part)) if html_part is not None else ""

    chosen = choose_best_body(plain_text, html_text)
    stripped = strip_quoted_sections(chosen)
    return normalize_whitespace(stripped)


def choose_best_body(plain_text: str, html_text: str) -> str:
    plain = normalize_whitespace(plain_text)
    html = normalize_whitespace(html_text)
    if plain and not looks_like_placeholder_plain(plain, html):
        return plain
    if html:
        return html
    return plain


def looks_like_placeholder_plain(plain_text: str, html_text: str) -> bool:
    if not plain_text or not html_text:
        return False
    lowered = plain_text.lower()
    markers = (
        "view this email in your browser",
        "please view this email in html",
        "this is an html message",
    )
    if any(marker in lowered for marker in markers):
        return True
    return len(plain_text) < 20 and len(html_text) > 200


def html_part_to_markdown(html: str) -> str:
    if not html.strip():
        return ""
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style", "head", "meta", "title", "noscript", "svg"]):
        tag.decompose()
    for hidden in soup.select('[style*="display:none"], [style*="display: none"]'):
        hidden.decompose()
    unwrap_marketing_layout_tables(soup)
    for img in soup.find_all("img"):
        alt_attr = img.get("alt") or img.get("title") or "inline image"
        alt = alt_attr if isinstance(alt_attr, str) else "inline image"
        src_attr = img.get("src", "")
        src = src_attr if isinstance(src_attr, str) else ""
        if src.startswith("cid:"):
            img.replace_with(f"【内联图片：{alt}，{src}】")
            continue
        if is_tracking_url(src):
            img.decompose()
            continue
        img.replace_with(f"【图片：{alt}】")
    for link in soup.find_all("a"):
        href_attr = link.get("href") or ""
        href = href_attr if isinstance(href_attr, str) else ""
        cleaned = simplify_tracking_url(href)
        if cleaned != href:
            link["href"] = cleaned
        text = link.get_text(" ", strip=True)
        if text and len(cleaned) > 120 and is_tracking_url(href):
            link.string = text
    text = html_to_markdown(str(soup), heading_style="ATX", bullets="-")
    return cleanup_markdown_noise(text)


def strip_quoted_sections(text: str) -> str:
    lines = text.splitlines()
    kept: list[str] = []
    for index, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith(">"):
            break
        if any(pattern.match(stripped) for pattern in QUOTE_SEPARATORS):
            break
        if index > 0 and stripped in {"From:", "Sent:", "To:", "Subject:"}:
            break
        kept.append(line)
    return "\n".join(kept).strip()


def normalize_whitespace(text: str) -> str:
    text = ZERO_WIDTH_CHARS.sub("", text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def collect_attachments(message: EmailMessage) -> list[AttachmentInfo]:
    attachments: list[AttachmentInfo] = []
    for part in message.iter_attachments():
        payload = part.get_payload(decode=True) or b""
        attachments.append(
            AttachmentInfo(
                filename=part.get_filename(),
                content_type=part.get_content_type(),
                size=len(payload),
                disposition=part.get_content_disposition(),
                content_id=part.get("Content-ID"),
            )
        )
    return attachments


def part_text(part: EmailMessage | Message | None) -> str:
    if part is None:
        return ""
    try:
        content = cast(Any, part).get_content()
        return content if isinstance(content, str) else str(content)
    except Exception:
        payload = cast(bytes, part.get_payload(decode=True) or b"")
        charset = part.get_content_charset() or "utf-8"
        return payload.decode(charset, errors="replace")


def header_value(message: Message, key: str) -> str | None:
    value = message.get(key)
    if value is None:
        return None
    return str(value).strip() or None


def select_body_parts(message: EmailMessage) -> tuple[EmailMessage | None, EmailMessage | None]:
    plain_part: EmailMessage | None = None
    html_part: EmailMessage | None = None
    for part in message.walk():
        if not isinstance(part, EmailMessage):
            continue
        if part.is_multipart():
            continue
        disposition = part.get_content_disposition()
        if disposition == "attachment":
            continue
        content_type = part.get_content_type()
        if content_type == "text/plain" and plain_part is None:
            plain_part = part
        if content_type == "text/html" and html_part is None:
            html_part = part
    return plain_part, html_part


def unwrap_marketing_layout_tables(soup: BeautifulSoup) -> None:
    for table in list(soup.find_all("table")):
        if table.parent is None:
            continue
        rows = table.find_all("tr")
        cells = table.find_all(["td", "th"])
        if not rows or not cells:
            if table.parent is not None:
                table.unwrap()
            continue
        text_cells = [cell.get_text(" ", strip=True) for cell in cells]
        non_empty = [value for value in text_cells if value]
        if not non_empty:
            table.decompose()
            continue
        empty_ratio = 1 - (len(non_empty) / max(len(text_cells), 1))
        if len(rows) > 8 or empty_ratio > 0.45:
            replacement_lines = [value for value in non_empty if len(value) > 2]
            replacement_text = "\n- " + "\n- ".join(dict.fromkeys(replacement_lines))
            table.replace_with(soup.new_string(replacement_text))


def cleanup_markdown_noise(text: str) -> str:
    lines = [line.rstrip() for line in text.splitlines()]
    cleaned: list[str] = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            cleaned.append("")
            continue
        if stripped == "-->":
            continue
        if stripped.startswith("![](") and is_tracking_url(stripped[4:-1]):
            continue
        if stripped.startswith("[http") and len(stripped) > 140:
            cleaned.append("[long tracking link omitted]")
            continue
        if stripped.count("|") >= 4 and stripped.replace("|", "").replace("-", "").strip() == "":
            continue
        cleaned.append(line)
    text = "\n".join(cleaned)
    text = ZERO_WIDTH_CHARS.sub("", text)
    text = re.sub(r"\n(\[[^\]]+\]\([^)]{120,}\)\n)+", "\n[long marketing links omitted]\n", text)
    return text


def is_tracking_url(url: str) -> bool:
    lowered = url.lower()
    return any(marker in lowered for marker in TRACKING_HOST_MARKERS)


def simplify_tracking_url(url: str) -> str:
    if not url:
        return url
    parsed = urllib.parse.urlparse(url)
    if not parsed.scheme or not parsed.netloc:
        return url
    if not is_tracking_url(url):
        return url
    query = urllib.parse.parse_qs(parsed.query)
    for key in ("url", "u", "target", "redirect", "redirect_url", "destination"):
        values = query.get(key)
        if values:
            candidate = urllib.parse.unquote(values[0])
            if candidate.startswith("http://") or candidate.startswith("https://"):
                return candidate
    return f"{parsed.scheme}://{parsed.netloc}{parsed.path}"
