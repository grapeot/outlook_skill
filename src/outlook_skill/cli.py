from __future__ import annotations

import argparse
import json
import sys
import time
from collections.abc import Sequence
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import Any, cast

from tqdm import tqdm

from .auth import AuthManager
from .calendar import create_calendar_invite, list_calendar_events
from .config import doctor_info, load_settings
from .downloader import download_recent_mail, list_local_mail, read_local_mail
from .errors import OutlookSkillError
from .exporter import export_local_to_markdown
from .markdown_renderer import render_maildir_to_markdown
from .replier import reply_to_message
from .sender import send_mail
from .spam_triage import add_rule, apply_rules, list_labels, list_rules


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="outlook-skill")
    subparsers = parser.add_subparsers(dest="command", required=True)

    doctor_parser = subparsers.add_parser("doctor")
    doctor_subparsers = doctor_parser.add_subparsers(dest="doctor_command", required=True)
    doctor_config = doctor_subparsers.add_parser("config")
    doctor_config.add_argument("--format", choices=("json", "text"), default="text")

    auth_parser = subparsers.add_parser("auth")
    auth_subparsers = auth_parser.add_subparsers(dest="auth_command", required=True)
    auth_login = auth_subparsers.add_parser("login")
    auth_login.add_argument("--device-code", action="store_true")
    auth_login.add_argument("--format", choices=("json", "text"), default="json")
    auth_status = auth_subparsers.add_parser("status")
    auth_status.add_argument("--format", choices=("json", "text"), default="json")

    mail_parser = subparsers.add_parser("mail")
    mail_subparsers = mail_parser.add_subparsers(dest="mail_command", required=True)

    mail_download = mail_subparsers.add_parser("download")
    mail_download.add_argument("--days", type=int, default=7)
    mail_download.add_argument("--limit", type=int, default=200)
    mail_download.add_argument("--folder", action="append")
    mail_download.add_argument("--format", choices=("json", "text"), default="json")

    mail_list_local = mail_subparsers.add_parser("list-local")
    mail_list_local.add_argument("--limit", type=int, default=50)
    mail_list_local.add_argument("--format", choices=("json", "text"), default="json")

    mail_read = mail_subparsers.add_parser("read")
    mail_read.add_argument("--subject", help="case-insensitive substring match against Subject")
    mail_read.add_argument("--from", dest="from_filter", help="case-insensitive substring match against From")
    mail_read.add_argument("--graph-id", dest="graph_id", help="exact Graph message id (skips subject/from matching)")
    mail_read.add_argument("--index", type=int, help="pick the Nth match (0-based, date-desc order)")
    mail_read.add_argument("--latest", action="store_true", help="pick the most recent match")
    mail_read.add_argument("--full", action="store_true", help="disable body truncation")
    mail_read.add_argument("--format", choices=("json", "text"), default="text")

    mail_export_md = mail_subparsers.add_parser("export-md")
    mail_export_md.add_argument("--days", type=int, default=None, help="only export messages received within N days")
    mail_export_md.add_argument("--folder", action="append", help="filter by folder (may repeat)")
    mail_export_md.add_argument("--subject", help="case-insensitive subject substring filter")
    mail_export_md.add_argument("--from", dest="from_filter", help="case-insensitive from substring filter")
    mail_export_md.add_argument("--output-dir", help="override output directory (default data/mail/markdown)")
    mail_export_md.add_argument("--force", action="store_true", help="overwrite existing markdown files")
    mail_export_md.add_argument("--format", choices=("json", "text"), default="text")

    for reply_command in ("reply", "reply-all"):
        mail_reply = mail_subparsers.add_parser(reply_command)
        mail_reply.add_argument("--graph-id", dest="graph_id", required=True, help="Graph message id of the email to reply to")
        mail_reply.add_argument("--body-file", dest="body_file", required=True, help="path to a file containing the reply body")
        mail_reply.add_argument("--body-format", choices=("text", "html", "markdown", "md"), default="text")
        mail_reply.add_argument("--attach", action="append", help="path to attach (may repeat)")
        mail_reply.add_argument("--to", action="append", help="override reply recipients (may repeat)")
        mail_reply.add_argument("--cc", action="append", help="override cc recipients (may repeat)")
        mail_reply.add_argument("--dry-run", action="store_true", help="create the draft but do not send")
        mail_reply.add_argument("--format", choices=("json", "text"), default="json")

    mail_send = mail_subparsers.add_parser("send")
    mail_send.add_argument("--to", action="append", required=True, help="recipient email address (may repeat)")
    mail_send.add_argument("--cc", action="append", help="cc recipient email address (may repeat)")
    mail_send.add_argument("--bcc", action="append", help="bcc recipient email address (may repeat)")
    mail_send.add_argument("--subject", required=True)
    mail_send.add_argument("--body-file", dest="body_file", required=True, help="path to a file containing the email body")
    mail_send.add_argument("--body-format", choices=("text", "html", "markdown", "md"), default="text")
    mail_send.add_argument("--attach", action="append", help="path to attach; currently supports files up to 3 MB (may repeat)")
    mail_send.add_argument("--no-save-to-sent-items", action="store_true")
    mail_send.add_argument("--dry-run", action="store_true", help="validate and render the payload but do not call Graph")
    mail_send.add_argument("--format", choices=("json", "text"), default="json")

    mail_render_markdown = mail_subparsers.add_parser("render-markdown")
    mail_render_markdown.add_argument("--input-dir")
    mail_render_markdown.add_argument("--output-dir")
    mail_render_markdown.add_argument("--limit", type=int)
    mail_render_markdown.add_argument("--sample-size", type=int)
    mail_render_markdown.add_argument("--seed", type=int, default=0)
    mail_render_markdown.add_argument("--workers", type=int, default=8)
    mail_render_markdown.add_argument("--format", choices=("json", "text"), default="json")

    calendar_parser = subparsers.add_parser("calendar")
    calendar_subparsers = calendar_parser.add_subparsers(dest="calendar_command", required=True)
    calendar_invite = calendar_subparsers.add_parser("invite")
    calendar_invite.add_argument("--to", action="append", help="attendee email address to invite (may repeat)")
    calendar_invite.add_argument("--optional-attendee", action="append", help="optional attendee email address (may repeat)")
    calendar_invite.add_argument("--subject", required=True)
    calendar_invite.add_argument("--start", required=True, help="local date-time, e.g. 2026-05-06T10:00:00")
    calendar_invite.add_argument("--end", required=True, help="local date-time, e.g. 2026-05-06T10:30:00")
    calendar_invite.add_argument("--timezone", default="UTC", help="Microsoft Graph timeZone value, e.g. UTC or Pacific Standard Time")
    calendar_invite.add_argument("--location")
    calendar_invite.add_argument("--body-file", dest="body_file", help="optional path to a file containing the event body")
    calendar_invite.add_argument("--body-format", choices=("text", "html", "markdown", "md"), default="text")
    calendar_invite.add_argument("--dry-run", action="store_true", help="validate and render the payload but do not call Graph")
    calendar_invite.add_argument("--format", choices=("json", "text"), default="json")

    calendar_list = calendar_subparsers.add_parser("list")
    calendar_list.add_argument("--start", help="start date (YYYY-MM-DD), defaults to today")
    calendar_list.add_argument("--end", help="end date (YYYY-MM-DD), defaults to today+60 days")
    calendar_list.add_argument("--skip-recurring", default="daily,weekly",
                              help="comma-separated recurrence types to skip, 'all' or 'none'. Default: daily,weekly")
    calendar_list.add_argument("--format", choices=("json", "text"), default="json")

    triage_parser = subparsers.add_parser("triage")
    triage_subparsers = triage_parser.add_subparsers(dest="triage_command", required=True)

    triage_add_rule = triage_subparsers.add_parser("add-rule")
    triage_add_rule.add_argument("--name", required=True)
    triage_add_rule.add_argument("--label", default="spam")
    triage_add_rule.add_argument("--type", required=True)
    triage_add_rule.add_argument("--pattern", required=True)
    triage_add_rule.add_argument("--confidence", type=float, default=0.8)
    triage_add_rule.add_argument("--format", choices=("json", "text"), default="json")

    triage_list_rules = triage_subparsers.add_parser("list-rules")
    triage_list_rules.add_argument("--format", choices=("json", "text"), default="json")

    triage_apply = triage_subparsers.add_parser("apply")
    triage_apply.add_argument("--folder")
    triage_apply.add_argument("--limit", type=int, default=100)
    triage_apply.add_argument("--label", default="spam")
    triage_apply.add_argument("--format", choices=("json", "text"), default="json")

    triage_list_labels = triage_subparsers.add_parser("list-labels")
    triage_list_labels.add_argument("--label")
    triage_list_labels.add_argument("--limit", type=int, default=100)
    triage_list_labels.add_argument("--format", choices=("json", "text"), default="json")

    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        if args.command == "doctor":
            payload = doctor_info()
            emit_output(payload, args.format)
            return 0

        settings = load_settings()

        if args.command == "auth":
            auth_manager = AuthManager(settings)
            if args.auth_command == "login":
                payload = auth_manager.login(device_code=args.device_code)
                emit_output(payload, args.format)
                return 0
            if args.auth_command == "status":
                payload = auth_manager.status()
                emit_output(payload, args.format)
                return 0

        if args.command == "mail":
            if args.mail_command == "download":
                payload = download_recent_mail(
                    settings,
                    days=args.days,
                    limit=args.limit,
                    folders=tuple(args.folder) if args.folder else None,
                    progress_callback=create_progress_callback(args.format),
                )
                emit_output(payload, args.format)
                return 0
            if args.mail_command == "list-local":
                payload = list_local_mail(settings, limit=args.limit)
                emit_output(payload, args.format)
                return 0
            if args.mail_command == "read":
                payload = read_local_mail(
                    settings,
                    subject=args.subject,
                    from_filter=args.from_filter,
                    graph_id=args.graph_id,
                    index=args.index,
                    latest=args.latest,
                    full=args.full,
                )
                emit_read_output(payload, args.format)
                return 0
            if args.mail_command == "export-md":
                payload = export_local_to_markdown(
                    settings,
                    days=args.days,
                    folders=tuple(args.folder) if args.folder else None,
                    subject=args.subject,
                    from_filter=args.from_filter,
                    force=args.force,
                    output_dir=Path(args.output_dir) if args.output_dir else None,
                )
                emit_output(payload, args.format)
                return 0
            if args.mail_command in ("reply", "reply-all"):
                attachments = tuple(Path(p) for p in (args.attach or ()))
                body_path = Path(args.body_file)
                if not body_path.exists():
                    raise OutlookSkillError(f"--body-file not found: {body_path}")
                body_text = body_path.read_text(encoding="utf-8")
                payload = reply_to_message(
                    settings,
                    graph_id=args.graph_id,
                    body_text=body_text,
                    body_format=args.body_format,
                    attachments=attachments,
                    to_override=tuple(args.to) if args.to else (),
                    cc_override=tuple(args.cc) if args.cc else (),
                    dry_run=args.dry_run,
                    reply_all=args.mail_command == "reply-all",
                )
                emit_output(payload, args.format)
                return 0
            if args.mail_command == "send":
                attachments = tuple(Path(p) for p in (args.attach or ()))
                body_path = Path(args.body_file)
                if not body_path.exists():
                    raise OutlookSkillError(f"--body-file not found: {body_path}")
                body_text = body_path.read_text(encoding="utf-8")
                payload = send_mail(
                    settings,
                    to=tuple(args.to),
                    subject=args.subject,
                    body_text=body_text,
                    body_format=args.body_format,
                    cc=tuple(args.cc) if args.cc else (),
                    bcc=tuple(args.bcc) if args.bcc else (),
                    attachments=attachments,
                    save_to_sent_items=not args.no_save_to_sent_items,
                    dry_run=args.dry_run,
                )
                emit_output(payload, args.format)
                return 0
            if args.mail_command == "render-markdown":
                payload = render_maildir_to_markdown(
                    settings,
                    input_dir=Path(args.input_dir) if args.input_dir else None,
                    output_dir=Path(args.output_dir) if args.output_dir else None,
                    limit=args.limit,
                    sample_size=args.sample_size,
                    seed=args.seed,
                    workers=args.workers,
                    progress_callback=create_render_progress_callback(args.format),
                )
                emit_output(payload, args.format)
                return 0

        if args.command == "calendar":
            if args.calendar_command == "invite":
                body_text = ""
                if args.body_file:
                    body_path = Path(args.body_file)
                    if not body_path.exists():
                        raise OutlookSkillError(f"--body-file not found: {body_path}")
                    body_text = body_path.read_text(encoding="utf-8")
                payload = create_calendar_invite(
                    settings,
                    subject=args.subject,
                    start=args.start,
                    end=args.end,
                    timezone=args.timezone,
                    attendees=tuple(args.to) if args.to else (),
                    body_text=body_text,
                    body_format=args.body_format,
                    location=args.location,
                    optional_attendees=tuple(args.optional_attendee) if args.optional_attendee else (),
                    dry_run=args.dry_run,
                )
                emit_output(payload, args.format)
                return 0
            if args.calendar_command == "list":
                start = args.start or datetime.now(UTC).strftime("%Y-%m-%d")
                end = args.end or (datetime.now(UTC) + timedelta(days=60)).strftime("%Y-%m-%d")
                payload = list_calendar_events(
                    settings,
                    start_date=start,
                    end_date=end,
                    skip_recurring=args.skip_recurring,
                )
                emit_output(payload, args.format)
                return 0

        if args.command == "triage":
            if args.triage_command == "add-rule":
                payload = add_rule(
                    settings,
                    rule_name=args.name,
                    label=args.label,
                    rule_type=args.type,
                    pattern=args.pattern,
                    confidence=args.confidence,
                )
                emit_output(payload, args.format)
                return 0
            if args.triage_command == "list-rules":
                payload = list_rules(settings)
                emit_output(payload, args.format)
                return 0
            if args.triage_command == "apply":
                payload = apply_rules(settings, folder=args.folder, limit=args.limit, label=args.label)
                emit_output(payload, args.format)
                return 0
            if args.triage_command == "list-labels":
                payload = list_labels(settings, label=args.label, limit=args.limit)
                emit_output(payload, args.format)
                return 0

        raise OutlookSkillError("Unsupported command.")
    except OutlookSkillError as exc:
        print(str(exc), file=sys.stderr)
        return 1


def emit_output(payload: object, format_name: str) -> None:
    if format_name == "json":
        print(json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True))
        return
    if isinstance(payload, dict):
        for key, value in payload.items():
            if isinstance(value, (dict, list)):
                value = json.dumps(value, ensure_ascii=False)
            print(f"{key}: {value}")
        return
    print(payload)


def emit_read_output(payload: dict[str, Any], format_name: str) -> None:
    if format_name == "json":
        print(json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True))
        return
    body = payload.get("body")
    if body is None:
        print(f"Matched {payload.get('match_count', 0)} messages; none selected.")
        note = payload.get("note")
        if note:
            print(f"({note})")
        print()
        for match in payload.get("matches", []) or []:
            print(
                f"[{match['index']}] {match.get('received_at') or '(no date)'}  "
                f"{match.get('from_addr') or ''}  |  {match.get('subject') or '(no subject)'}  "
                f"[{match.get('folder') or ''}] graph_id={match.get('graph_id')}"
            )
        return
    print(f"From: {payload.get('from') or ''}")
    print(f"To: {payload.get('to') or ''}")
    cc = payload.get("cc")
    if cc:
        print(f"Cc: {cc}")
    print(f"Subject: {payload.get('subject') or ''}")
    print(f"Date: {payload.get('date') or ''}")
    print(f"Folder: {payload.get('folder') or ''}")
    print(f"Graph-ID: {payload.get('graph_id') or ''}")
    print(f"Body-Source: {payload.get('body_source') or ''}")
    print()
    print(body)
    if payload.get("body_truncated"):
        remaining = payload.get("body_truncated_remaining_chars") or 0
        print(f"\n[truncated, {remaining} more chars; rerun with --full to see all]")


def create_progress_callback(format_name: str):
    stream = sys.stderr if format_name == "json" else sys.stdout
    progress_bar: Any = None
    last_current = 0
    started_at = time.monotonic()

    def report(*, current: int, total: int, downloaded: int, skipped: int) -> None:
        nonlocal progress_bar, last_current
        if progress_bar is None:
            progress_bar = tqdm(
                total=total,
                desc="Downloading",
                unit="msg",
                file=stream,
                dynamic_ncols=True,
                leave=True,
            )
        delta = max(current - last_current, 0)
        progress = cast(Any, progress_bar)
        if delta:
            progress.update(delta)
        elapsed = max(time.monotonic() - started_at, 0.001)
        rate = current / elapsed if current else 0.0
        remaining = max(total - current, 0)
        eta_seconds = int(remaining / rate) if rate > 0 else 0
        progress.set_postfix(downloaded=downloaded, skipped=skipped, eta=format_eta(eta_seconds))
        last_current = current
        if total == 0 or current >= total:
            progress.close()

    return report


def create_render_progress_callback(format_name: str):
    stream = sys.stderr if format_name == "json" else sys.stdout
    progress_bar: Any = None
    last_current = 0
    started_at = time.monotonic()

    def report(*, current: int, total: int, rendered: int, skipped: int) -> None:
        nonlocal progress_bar, last_current
        if progress_bar is None:
            progress_bar = tqdm(
                total=total,
                desc="Rendering MD",
                unit="msg",
                file=stream,
                dynamic_ncols=True,
                leave=True,
            )
        delta = max(current - last_current, 0)
        progress = cast(Any, progress_bar)
        if delta:
            progress.update(delta)
        elapsed = max(time.monotonic() - started_at, 0.001)
        rate = current / elapsed if current else 0.0
        remaining = max(total - current, 0)
        eta_seconds = int(remaining / rate) if rate > 0 else 0
        progress.set_postfix(rendered=rendered, skipped=skipped, eta=format_eta(eta_seconds))
        last_current = current
        if total == 0 or current >= total:
            progress.close()

    return report


def format_eta(seconds: int) -> str:
    if seconds < 60:
        return f"{seconds}s"
    minutes, secs = divmod(seconds, 60)
    if minutes < 60:
        return f"{minutes}m{secs:02d}s"
    hours, minutes = divmod(minutes, 60)
    return f"{hours}h{minutes:02d}m"


if __name__ == "__main__":
    raise SystemExit(main())
