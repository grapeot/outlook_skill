# Outlook Skill — Agent Reference

## When to use

- Download Outlook.com email to local storage
- Search or read locally cached email
- Render `.eml` to AI-oriented Markdown
- Create standalone drafts, send email, reply in-thread, or reply-all in-thread via Microsoft Graph
- Create calendar invites or list upcoming events
- Apply rule-based triage labels to local messages

## Prerequisites

- Working directory: project root (alongside `pyproject.toml`)
- Python: `.venv/` (created with `uv`)
- `.env` configured with `OUTLOOK_EMAIL`, `OUTLOOK_CLIENT_ID`, `OUTLOOK_AUTHORITY`
- Completed `auth login` with valid token cache at `data/mail/oauth_token_cache.json`

## Commands

All commands run from the project root.

```bash
.venv/bin/python -m outlook_skill.cli doctor config --format json
.venv/bin/python -m outlook_skill.cli auth login
.venv/bin/python -m outlook_skill.cli auth status --format json

.venv/bin/python -m outlook_skill.cli mail download --days 7 --limit 200 --format json
.venv/bin/python -m outlook_skill.cli mail download --days 7 --limit 200 --folder Inbox --folder Archive --format json

.venv/bin/python -m outlook_skill.cli mail list-local --limit 50 --format json
.venv/bin/python -m outlook_skill.cli mail read --subject "<substring>" [--from <substring>] [--latest | --index N | --graph-id <id>] [--full]
.venv/bin/python -m outlook_skill.cli mail export-md [--days N] [--folder <name>]... [--subject <substr>] [--from <substr>] [--force]

.venv/bin/python -m outlook_skill.cli mail draft --subject <subject> --body-file <path> [--body-format text|html|markdown] [--to <addr>]... [--cc <addr>]... [--bcc <addr>]... [--attach <path>]...
.venv/bin/python -m outlook_skill.cli mail reply --graph-id <id> --body-file <path> [--body-format text|html|markdown] [--attach <path>]... [--to <addr>]... [--cc <addr>]... [--dry-run]
.venv/bin/python -m outlook_skill.cli mail reply-all --graph-id <id> --body-file <path> [--body-format text|html|markdown] [--attach <path>]... [--to <addr>]... [--cc <addr>]... [--dry-run]
.venv/bin/python -m outlook_skill.cli mail send --to <addr> --subject <subject> --body-file <path> [--body-format text|html|markdown] [--cc <addr>]... [--bcc <addr>]... [--attach <path>]... [--dry-run]

.venv/bin/python -m outlook_skill.cli calendar invite [--to <addr>]... --subject <subject> --start <YYYY-MM-DDTHH:MM:SS> --end <YYYY-MM-DDTHH:MM:SS> [--timezone UTC] [--optional-attendee <addr>]... [--location <text>] [--body-file <path>] [--dry-run]
.venv/bin/python -m outlook_skill.cli calendar list [--start YYYY-MM-DD] [--end YYYY-MM-DD] [--skip-recurring daily,weekly] --format json
```

A convenience wrapper is also available:

```bash
scripts/outlook mail download --days 7 --limit 200 --format json
```

### `mail download`

Pulls recent mail from Microsoft Graph. Default folders: `Inbox` + `Archive`. Explicit `--folder` overrides the defaults. Each message is saved as raw `.eml` with a SQLite metadata record. The idempotency key is `(account, folder, uidvalidity="graph", uid=graph_message_id)` — re-running on the same time window skips already-downloaded messages.

### `mail read`

Looks up locally cached messages by subject substring or Graph message ID. Queries SQLite; reads body from `.eml` (plain text preferred, HTML as fallback). Default body truncation at 10,000 characters; `--full` disables truncation. When multiple messages match, lists them with indices; use `--latest`, `--index N`, or `--graph-id` to select one.

### `mail export-md`

Exports locally cached `.eml` to YAML frontmatter Markdown in `data/mail/markdown/`. Filenames: `{date}_{folder}_{subject-slug}_{short-graph-id}.md`. Optional filters: `--days N`, `--folder`, `--subject`, `--from`. `--force` overwrites existing files.

### `mail send`

Sends standalone email via `POST /me/sendMail`. Requires `Mail.Send` scope. Inline attachments up to 3 MB. `--body-format markdown` converts Markdown to HTML before sending. `--dry-run` validates the payload without calling Graph.

### `mail draft`

Creates a standalone draft via `POST /me/messages`. Requires `Mail.ReadWrite` scope. `--to`, `--cc`, and `--bcc` are optional, so agents can save no-recipient drafts for human review. The command never sends; JSON output includes `draft_id`, `web_link`, recipient summaries, and `sent: false`. Inline attachments up to 3 MB are supported.

### `mail reply` / `mail reply-all`

Replies in-thread via Graph draft flow. `mail reply` uses `createReply`; `mail reply-all` uses `createReplyAll`. Both require `Mail.ReadWrite` + `Mail.Send`. Attachments ≤3 MB use inline `fileAttachment`; larger files use upload session. `--to` / `--cc` override the draft recipients after Graph creates the draft. `--dry-run` creates the draft without sending. JSON output includes `operation: reply` or `operation: reply_all`.

Graph's `createReply`/`createReplyAll` automatically includes the original message as a quoted thread in the draft body. The replier preserves this quoted content and prepends the user's reply above it. When content types differ (user sends Text, Graph returns HTML), the text side is converted to HTML for consistency.

**Default reply mode:** When a user asks to "reply" or "回信" without specifying, use `mail reply-all` to preserve all original recipients. Use `mail reply` (reply to sender only) only when the user explicitly requests it.

### `calendar invite`

Creates a calendar event via `POST /me/calendar/events`. Requires `Calendars.ReadWrite`. By default it creates an attendee-less appointment; pass `--to` or `--optional-attendee` only when someone should be invited. Supports location and body. `--dry-run` validates the payload without calling Graph.

### `calendar list`

Reads upcoming events from the default calendar via `GET /me/calendar/events`. Default: today to today+60 days. `--skip-recurring daily,weekly` (default) filters out daily and weekly recurring events; `all` filters all recurring; `none` shows everything.

## AI workflow

The standard pipeline:

1. `mail download` — pull recent `.eml` to local storage
2. `mail export-md` — generate YAML frontmatter Markdown
3. Use `grep` or text search on `data/mail/markdown/` for retrieval
4. `mail read --graph-id <id>` — read full body of a specific message
5. `mail draft`, `mail reply --graph-id <id>`, `mail reply-all --graph-id <id>`, or `mail send` — compose responses

## Defaults

- `mail download` targets `Inbox` + `Archive`
- `--folder` overrides the default list
- `--format json`: results to stdout, progress to stderr
- Idempotency: `(account, folder, graph, message_id)` — same message in same folder is skipped on re-download
- Same message in different folders produces separate local records

## Output contract

`mail download --format json` returns:

```json
{
  "folders": ["Inbox", "Archive"],
  "days": 7,
  "downloaded_count": 42,
  "skipped_existing_count": 158,
  "messages": [
    {
      "folder": "Inbox",
      "graph_id": "AAMk...",
      "subject": "Meeting tomorrow",
      "from_addr": "colleague@example.com",
      "received_at": "2026-05-10T14:30:00Z",
      "mime_path": "data/mail/messages/..."
    }
  ]
}
```

`mail export-md` writes files to `data/mail/markdown/` with YAML frontmatter containing `graph_id`, `folder`, `from`, `to`, `cc`, `subject`, `date`, and `body_source`.

## Storage

| Layer | Path |
|---|---|
| Raw MIME | `data/mail/messages/*.eml` |
| Markdown (canonical YAML frontmatter) | `data/mail/markdown/*.md` |
| SQLite | `data/mail/mail.db` |
| Token cache | `data/mail/oauth_token_cache.json` |

## Known caveats

- Idempotency uses Graph message ID. This is stable enough for recent-time-window download but may change if messages move between folders on the server. Long-term synchronization across folder moves would need immutable ID evaluation.
- Progress bars write to stderr in JSON mode; stdout contains only the valid JSON result.
- The system does not extract attachment content, sync folder state, or modify server-side message flags.
- Markdown rendering optimizes for AI consumption, not human readability. It strips quotes, tracking infrastructure, and layout tables.

## Acceptance criteria

1. `doctor config` reports `client_id_configured=true`
2. `auth status` reports `auth_ready=true`
3. `mail download` returns valid JSON and writes `.eml` files
4. Re-running on the same folder/time window increases `skipped_existing_count` rather than duplicating writes
5. `mail export-md` writes canonical `.md` files with clean stdout in JSON mode
