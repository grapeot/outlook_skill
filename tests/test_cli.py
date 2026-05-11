import json

from outlook_skill import cli


def test_doctor_config_json_output(monkeypatch, capsys):
    monkeypatch.setattr(cli, "doctor_info", lambda: {"client_id_configured": True})

    exit_code = cli.main(["doctor", "config", "--format", "json"])

    captured = capsys.readouterr()
    assert exit_code == 0
    assert json.loads(captured.out)["client_id_configured"] is True


def test_auth_status_uses_auth_manager(monkeypatch, capsys):
    class FakeAuthManager:
        def __init__(self, settings):
            self.settings = settings

        def status(self):
            return {"auth_ready": True}

    monkeypatch.setattr(cli, "load_settings", lambda: object())
    monkeypatch.setattr(cli, "AuthManager", FakeAuthManager)

    exit_code = cli.main(["auth", "status", "--format", "json"])

    captured = capsys.readouterr()
    assert exit_code == 0
    assert json.loads(captured.out)["auth_ready"] is True


def test_mail_list_local_uses_store(monkeypatch, capsys):
    monkeypatch.setattr(cli, "load_settings", lambda: object())
    monkeypatch.setattr(cli, "list_local_mail", lambda settings, limit: {"count": 1, "messages": [{"uid": 1}]})

    exit_code = cli.main(["mail", "list-local", "--limit", "10", "--format", "json"])

    captured = capsys.readouterr()
    assert exit_code == 0
    assert json.loads(captured.out)["count"] == 1


def test_mail_download_keeps_json_stdout_clean_and_progress_on_stderr(monkeypatch, capsys):
    def fake_download_recent_mail(settings, *, days, limit, folders, progress_callback):
        progress_callback(current=0, total=2, downloaded=0, skipped=0)
        progress_callback(current=1, total=2, downloaded=1, skipped=0)
        progress_callback(current=2, total=2, downloaded=1, skipped=1)
        return {"downloaded_count": 1, "skipped_existing_count": 1, "messages": []}

    monkeypatch.setattr(cli, "load_settings", lambda: object())
    monkeypatch.setattr(cli, "download_recent_mail", fake_download_recent_mail)

    exit_code = cli.main(["mail", "download", "--days", "7", "--limit", "2", "--format", "json"])

    captured = capsys.readouterr()
    assert exit_code == 0
    assert json.loads(captured.out)["downloaded_count"] == 1
    assert "Downloading" in captured.err
    assert "downloaded=1" in captured.err
    assert "skipped=1" in captured.err


def test_mail_download_accepts_multiple_folder_flags(monkeypatch, capsys):
    captured_call = {}

    def fake_download_recent_mail(settings, *, days, limit, folders, progress_callback):
        captured_call["folders"] = folders
        return {"downloaded_count": 0, "skipped_existing_count": 0, "messages": [], "folders": list(folders or [])}

    monkeypatch.setattr(cli, "load_settings", lambda: object())
    monkeypatch.setattr(cli, "download_recent_mail", fake_download_recent_mail)

    exit_code = cli.main([
        "mail",
        "download",
        "--folder",
        "Inbox",
        "--folder",
        "Archive",
        "--format",
        "json",
    ])

    captured = capsys.readouterr()
    assert exit_code == 0
    assert captured_call["folders"] == ("Inbox", "Archive")
    assert json.loads(captured.out)["folders"] == ["Inbox", "Archive"]


def test_mail_render_markdown_passes_directories_and_sample(monkeypatch, capsys, tmp_path):
    captured_call = {}

    def fake_render_maildir_to_markdown(settings, *, input_dir, output_dir, limit, sample_size, seed, workers, progress_callback):
        captured_call["input_dir"] = input_dir
        captured_call["output_dir"] = output_dir
        captured_call["sample_size"] = sample_size
        progress_callback(current=0, total=1, rendered=0, skipped=0)
        progress_callback(current=1, total=1, rendered=1, skipped=0)
        return {"rendered_count": 1, "skipped_count": 0, "files": []}

    monkeypatch.setattr(cli, "load_settings", lambda: object())
    monkeypatch.setattr(cli, "render_maildir_to_markdown", fake_render_maildir_to_markdown)

    exit_code = cli.main([
        "mail",
        "render-markdown",
        "--input-dir",
        str(tmp_path / "messages"),
        "--output-dir",
        str(tmp_path / "markdowns"),
        "--sample-size",
        "10",
        "--format",
        "json",
    ])

    captured = capsys.readouterr()
    assert exit_code == 0
    assert captured_call["input_dir"] == tmp_path / "messages"
    assert captured_call["output_dir"] == tmp_path / "markdowns"
    assert captured_call["sample_size"] == 10
    assert json.loads(captured.out)["rendered_count"] == 1
    assert "Rendering MD" in captured.err


def test_triage_add_rule_cli(monkeypatch, capsys):
    monkeypatch.setattr(cli, "load_settings", lambda: object())
    monkeypatch.setattr(
        cli,
        "add_rule",
        lambda settings, rule_name, label, rule_type, pattern, confidence: {
            "rule_name": rule_name,
            "label": label,
            "rule_type": rule_type,
            "pattern": pattern,
            "confidence": confidence,
        },
    )

    exit_code = cli.main([
        "triage",
        "add-rule",
        "--name",
        "promo",
        "--label",
        "low_value",
        "--type",
        "sender_domain",
        "--pattern",
        "promo.example.com",
        "--format",
        "json",
    ])

    captured = capsys.readouterr()
    assert exit_code == 0
    assert json.loads(captured.out)["rule_type"] == "sender_domain"
    assert json.loads(captured.out)["label"] == "low_value"



def test_mail_send_cli_delegates_to_sender(monkeypatch, capsys, tmp_path):
    body = tmp_path / "body.md"
    body.write_text("Hello", encoding="utf-8")
    captured_call = {}

    def fake_send_mail(settings, *, to, subject, body_text, body_format, cc, bcc, attachments, save_to_sent_items, dry_run):
        captured_call.update({
            "to": to,
            "subject": subject,
            "body_text": body_text,
            "body_format": body_format,
            "cc": cc,
            "bcc": bcc,
            "attachments": attachments,
            "save_to_sent_items": save_to_sent_items,
            "dry_run": dry_run,
        })
        return {"sent": False, "dry_run": dry_run}

    monkeypatch.setattr(cli, "load_settings", lambda: object())
    monkeypatch.setattr(cli, "send_mail", fake_send_mail)

    exit_code = cli.main([
        "mail",
        "send",
        "--to",
        "duck@example.com",
        "--cc",
        "cc@example.com",
        "--bcc",
        "bcc@example.com",
        "--subject",
        "Hello",
        "--body-file",
        str(body),
        "--body-format",
        "markdown",
        "--no-save-to-sent-items",
        "--dry-run",
        "--format",
        "json",
    ])

    captured = capsys.readouterr()
    assert exit_code == 0
    assert json.loads(captured.out)["dry_run"] is True
    assert captured_call["to"] == ("duck@example.com",)
    assert captured_call["cc"] == ("cc@example.com",)
    assert captured_call["bcc"] == ("bcc@example.com",)
    assert captured_call["subject"] == "Hello"
    assert captured_call["body_text"] == "Hello"
    assert captured_call["body_format"] == "markdown"
    assert captured_call["save_to_sent_items"] is False


def test_calendar_invite_cli_delegates_to_calendar(monkeypatch, capsys, tmp_path):
    body = tmp_path / "body.txt"
    body.write_text("Agenda", encoding="utf-8")
    captured_call = {}

    def fake_create_calendar_invite(settings, *, subject, start, end, timezone, attendees, body_text, body_format, location, optional_attendees, dry_run):
        captured_call.update({
            "subject": subject,
            "start": start,
            "end": end,
            "timezone": timezone,
            "attendees": attendees,
            "body_text": body_text,
            "body_format": body_format,
            "location": location,
            "optional_attendees": optional_attendees,
            "dry_run": dry_run,
        })
        return {"created": False, "dry_run": dry_run}

    monkeypatch.setattr(cli, "load_settings", lambda: object())
    monkeypatch.setattr(cli, "create_calendar_invite", fake_create_calendar_invite)

    exit_code = cli.main([
        "calendar",
        "invite",
        "--to",
        "required@example.com",
        "--optional-attendee",
        "optional@example.com",
        "--subject",
        "Meeting",
        "--start",
        "2026-05-06T10:00:00",
        "--end",
        "2026-05-06T10:30:00",
        "--timezone",
        "Pacific Standard Time",
        "--location",
        "Zoom",
        "--body-file",
        str(body),
        "--body-format",
        "text",
        "--dry-run",
        "--format",
        "json",
    ])

    captured = capsys.readouterr()
    assert exit_code == 0
    assert json.loads(captured.out)["dry_run"] is True
    assert captured_call["attendees"] == ("required@example.com",)
    assert captured_call["optional_attendees"] == ("optional@example.com",)
    assert captured_call["subject"] == "Meeting"
    assert captured_call["timezone"] == "Pacific Standard Time"
    assert captured_call["location"] == "Zoom"
    assert captured_call["body_text"] == "Agenda"


def test_calendar_list_cli_delegates_to_calendar(monkeypatch, capsys):
    captured_call = {}

    def fake_list_calendar_events(settings, *, start_date, end_date, skip_recurring):
        captured_call.update({
            "start_date": start_date,
            "end_date": end_date,
            "skip_recurring": skip_recurring,
        })
        return {
            "start_date": start_date,
            "end_date": end_date,
            "total_count": 5,
            "shown_count": 3,
            "skipped_recurring": 2,
            "events": [
                {"subject": "Non-recurring meeting", "start": "2026-05-15T10:00:00"},
            ],
        }

    monkeypatch.setattr(cli, "load_settings", lambda: object())
    monkeypatch.setattr(cli, "list_calendar_events", fake_list_calendar_events)

    exit_code = cli.main([
        "calendar",
        "list",
        "--start",
        "2026-05-08",
        "--end",
        "2026-07-08",
        "--skip-recurring",
        "daily,weekly",
        "--format",
        "json",
    ])

    captured = capsys.readouterr()
    assert exit_code == 0
    assert json.loads(captured.out)["shown_count"] == 3
    assert captured_call["start_date"] == "2026-05-08"
    assert captured_call["end_date"] == "2026-07-08"
    assert captured_call["skip_recurring"] == "daily,weekly"
