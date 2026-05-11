# Outlook Skill

## What this repo is

An AI-first local pipeline for Outlook.com email. It uses Microsoft Graph to download mail as raw MIME, stores metadata in SQLite, and renders `.eml` into AI-oriented Markdown. The library, CLI, and skill document share a single contract — both humans and AI agents consume the same interface.

This is not an email client, a Graph debugging tool, or a background sync service. Write capabilities are constrained to three explicit operations: standalone send, in-thread reply, and calendar invite creation.

## Working environment

Use the project `.venv`. Python dependencies are managed with `uv`. Do not hardcode credentials; if `.env` uses `op://` references for 1Password, resolve them externally before passing values to the library.

Core commands:

```bash
python -m outlook_skill.cli doctor config
python -m outlook_skill.cli auth login
python -m outlook_skill.cli auth status
python -m outlook_skill.cli mail download --days 7 --limit 200 --format json
python -m outlook_skill.cli mail download --days 7 --limit 200 --folder Inbox --folder Archive --format json
python -m outlook_skill.cli mail render-markdown --input-dir data/mail/messages --output-dir data/mail/markdowns --workers 8 --format json
python -m outlook_skill.cli mail list-local --limit 50 --format json
python -m outlook_skill.cli mail send --to someone@example.com --subject "Subject" --body-file body.md --dry-run --format json
python -m outlook_skill.cli calendar invite --to someone@example.com --subject "Meeting" --start 2026-05-06T10:00:00 --end 2026-05-06T10:30:00 --dry-run --format json
```

## Code boundaries

`src/outlook_skill/` is the sole logic layer. Configuration, OAuth2, Graph transport, MIME storage, Markdown rendering, and local metadata all live within the package. The CLI handles argument parsing, environment assembly, library calls, and output formatting — nothing else.

Do not expand the system into arbitrary Graph passthrough. Each new capability must justify itself against the existing user stories.

## Security

Do not hardcode real client IDs, refresh tokens, or access tokens in code, tests, or docs. Downloaded `.eml` files, Markdown, SQLite metadata, and the token cache are sensitive — they stay out of git. Live integration tests default to skip; they require explicit opt-in via `OUTLOOK_ENABLE_LIVE_TESTS=1`. Write operations require additional flags: `OUTLOOK_LIVE_ALLOW_SEND=1` and `OUTLOOK_LIVE_ALLOW_CALENDAR_INVITE=1`.

## Testing and docs

After modifying the library or CLI, run `pytest` at minimum. Changes to OAuth2, Graph transport, Markdown rendering, or output contracts should also pass `doctor config` and a CLI smoke check. Update `docs/working.md` for significant changes — especially app registration prerequisites, token cache location, and Graph permission requirements.
