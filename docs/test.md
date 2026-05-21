# Test Strategy

## Scope

The test suite verifies five properties: OAuth2 configuration contracts are clear; the library and thin CLI produce consistent download output; `.eml` → Markdown rendering produces AI-usable output; standalone send and calendar invite payloads follow stable contracts; and real Outlook.com Graph writes are never triggered without explicit opt-in.

## Unit tests

Unit tests cover pure logic with no network dependency:

- Configuration loading and default value handling
- Graph client parameter construction for message listing and MIME download
- Token cache state detection
- Email metadata extraction from parsed MIME
- Plain text vs. HTML body selection, HTML-to-Markdown conversion, and noise cleaning
- Quote section stripping
- Attachment list rendering
- SQLite idempotent writes
- CLI argument parsing and downstream service call parameters
- `mail send`: Graph `/me/sendMail` payload construction, body format handling, recipient fields, attachment encoding, and dry-run behavior
- `mail reply` / `mail reply-all`: Graph `createReply` vs. `createReplyAll` endpoint selection, draft patching, attachment upload, operation output, and dry-run no-send behavior
- `calendar invite`: Graph `/me/calendar/events` payload construction, required vs. optional attendees, time validation, and dry-run behavior
- `calendar list`: Graph `/me/calendar/events` query parameter construction, time range filtering, daily/weekly recurring event filtering, response parsing, and pagination

## Mocked integration tests

Mocked integration tests verify the seams between library and CLI using fake auth managers, fake Graph clients, or monkeypatch:

- `doctor config` correctly reports missing client ID and token cache state
- `auth status` correctly reflects cache readiness
- `mail download` produces a clear error when not authenticated
- `mail render-markdown` writes `.md` files and keeps JSON stdout clean
- `mail list-local` outputs a stable JSON structure
- `mail send --dry-run` does not call Graph and outputs stable JSON
- `mail reply-all --dry-run` creates a Graph reply-all draft and does not send it
- `calendar invite --dry-run` does not call Graph and outputs stable JSON
- `calendar list` correctly parses Graph responses, filters daily/weekly recurring events, and handles empty results

## Live integration tests

Live integration tests exercise the full OAuth2 → Graph → local disk pipeline against a real Outlook.com account. They verify the end-to-end chain remains functional:

- Real `auth login` completes successfully
- Graph lists and downloads Outlook.com messages
- MIME `/$value` endpoint returns valid RFC 822 content
- `.eml` files and SQLite records are written to disk
- Real `.eml` files are batch-rendered to `.md`
- With `OUTLOOK_LIVE_ALLOW_SEND=1`, a test email is sent to self and observed in recent messages
- With `OUTLOOK_LIVE_ALLOW_CALENDAR_INVITE=1`, a future calendar invite is created for self
- `calendar list` reads events from the default calendar, confirms stable output structure, and correctly distinguishes recurring events

This layer defaults to skip. It runs only when `OUTLOOK_ENABLE_LIVE_TESTS=1` is set and a valid OAuth2 configuration is present. Write operations require additional allow flags: `OUTLOOK_LIVE_ALLOW_SEND=1` for email sending, `OUTLOOK_LIVE_ALLOW_CALENDAR_INVITE=1` for calendar invites. No test sends real email or creates real calendar events without explicit, separate opt-in.

## Running tests

Default (unit + mocked integration):

```bash
.venv/bin/python -m pytest -v
```

Live integration (requires valid OAuth2 config):

```bash
OUTLOOK_ENABLE_LIVE_TESTS=1 .venv/bin/python -m pytest -v -m live_integration

OUTLOOK_ENABLE_LIVE_TESTS=1 OUTLOOK_LIVE_ALLOW_SEND=1 OUTLOOK_LIVE_TEST_TO=your_account@outlook.com .venv/bin/python -m pytest -v -m live_integration tests/test_live_integration.py::test_live_send_mail_to_self

OUTLOOK_ENABLE_LIVE_TESTS=1 OUTLOOK_LIVE_ALLOW_CALENDAR_INVITE=1 OUTLOOK_LIVE_TEST_CALENDAR_TO=your_account@outlook.com .venv/bin/python -m pytest -v -m live_integration tests/test_live_integration.py::test_live_calendar_invite_to_self

OUTLOOK_ENABLE_LIVE_TESTS=1 .venv/bin/python -m pytest -v -m live_integration tests/test_live_integration.py::test_live_calendar_list
```

## Smoke checks

After significant changes, run at minimum:

```bash
.venv/bin/python -m outlook_skill.cli doctor config
.venv/bin/python -m outlook_skill.cli auth status --format json
.venv/bin/python -m outlook_skill.cli mail list-local --limit 20 --format json
.venv/bin/python -m outlook_skill.cli mail render-markdown --input-dir data/mail/messages --output-dir data/mail/markdowns --limit 20 --format json
.venv/bin/python -m outlook_skill.cli mail send --to your_account@outlook.com --subject "Dry run" --body-file body.md --dry-run --format json
.venv/bin/python -m outlook_skill.cli mail reply-all --graph-id <message_id> --body-file body.md --dry-run --format json
.venv/bin/python -m outlook_skill.cli calendar invite --to your_account@outlook.com --subject "Dry run" --start 2026-05-06T10:00:00 --end 2026-05-06T10:30:00 --dry-run --format json
.venv/bin/python -m outlook_skill.cli calendar list --start 2026-05-08 --end 2026-05-15 --format json
```

If no valid OAuth2 configuration is available, run only the default test suite — skip live downloads.
