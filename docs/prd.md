# Outlook Skill — Product Description

## What this is

`outlook_skill` is an AI-first local pipeline for Outlook.com email. It connects to Microsoft Graph, downloads raw MIME, stores lightweight metadata in SQLite, and renders `.eml` into AI-oriented Markdown. The library, CLI, and skill document share a single contract.

It is not an email client. It is not a Graph debugging tool. It is a constrained entry layer: a small set of stable, composable capabilities that turn a personal Outlook.com mailbox into a local, AI-consumable data source.

## What it does

The system provides ten capabilities, each with a clear module boundary:

1. **Configuration** (`config.py`): environment variables, OAuth2 authority, data directory, token cache path — all resolved at startup from `.env`
2. **Authentication** (`auth.py`): MSAL public client, interactive browser login, device code fallback, silent refresh, serialized token cache
3. **Graph transport** (`graph_client.py`): folder resolution (well-known names + custom folders), paginated message listing with `receivedDateTime` filtering, MIME `$value` download with automatic 401 retry
4. **Download orchestration** (`downloader.py`): iterates folders, applies idempotency check (`account + folder + uidvalidity + uid`), fetches MIME, extracts headers, persists to disk and SQLite, reports progress
5. **Local storage** (`store.py`): SQLite with three tables (`messages`, `spam_rules`, `triage_labels`), schema migration, unique constraint on `(account, folder, uidvalidity, uid)`, SHA-256 hashing, sanitized file paths
6. **Markdown rendering** (`markdown_renderer.py`): parallel `.eml`→`.md` conversion with configurable workers, plain-text preference with HTML fallback, quote stripping, zero-width character removal, tracking URL simplification, marketing layout table unwrapping
7. **Mail export** (`exporter.py`): YAML frontmatter Markdown export from local SQLite + `.eml`, with folder/subject/date filtering and slug-based filenames
8. **Standalone send** (`sender.py`): `POST /me/sendMail` with body-format conversion (Markdown→HTML), inline attachments ≤3MB, `saveToSentItems` toggle, dry-run mode
9. **In-thread reply** (`replier.py`): `createReply` / `createReplyAll` draft flow, attachment upload (inline for ≤3MB, upload session for larger), recipient override, dry-run that creates draft without sending
10. **Calendar** (`calendar.py`): `POST /me/calendar/events` for creating invites with required/optional attendees, `GET /me/calendar/events` for listing events with recurring-event filtering

A rule-based triage system (`spam_triage.py`) overlays the local message store with deterministic sender/subject/body matching, multi-label classification, and persistent labeling. Current rules (`spam` + `low_value`) cover approximately 29% of a 38,000-message corpus through three rounds of iterative refinement.

The CLI (`cli.py`) is a thin argparse shell that delegates to library functions, separating progress output (stderr) from machine-readable results (stdout) in JSON mode.

## Who uses it

Three user types, served by the same library contract:

**AI agents** are the primary consumers. They need stable function interfaces, predictable JSON output, explicit read/write boundaries, and safety defaults that prevent accidental state mutation. They should be able to download, render, search, and optionally send mail without understanding OAuth2 internals or Graph API details.

**Human operators** use the CLI for quick actions: first-time OAuth login, pulling recent Inbox + Archive mail, checking download counts, rendering `.eml` to `.md`, reading a specific message, or dry-running send/reply actions before committing.

**Project maintainers** need standard project structure, independent docs, layered tests, and a security model that keeps credential handling out of library code.

## Design principles

### Library-first, not script-first

A single `sync_outlook.py` script would couple configuration, OAuth2, Graph calls, MIME storage, Markdown rendering, and terminal output into one file. That works for a weekend prototype. For long-term use — and for AI agents that need to call individual pieces of the pipeline — the cost of that coupling compounds quickly. Every new feature pushes the script further past its breaking point. Test granularity is forced to "run the whole thing and check stdout."

Instead, the library exposes ten modules with clear contracts. The CLI is a consumer of the library, not the implementation. Tests can verify auth independently of transport, storage independently of rendering, and rules independently of download logic.

### Raw MIME is the source of truth

Attachments, headers, encoding details, and body parsing can all be redone later. Raw MIME, once lost, cannot be reconstructed. The system writes `.eml` to disk first, extracts metadata second, and renders Markdown third. SQLite records metadata for fast queries; `.eml` files are the canonical copy.

This ordering is not about storage efficiency. It is about irreversibility: a premature transformation that drops headers or encoding information cannot be undone from downstream data. The raw layer preserves the option to re-render, re-index, or apply future analysis techniques without re-downloading from Microsoft's servers.

### Markdown for AI, not for humans

The Markdown renderer optimizes for machine consumption: consistent structure, normalized whitespace, stripped tracking URLs, collapsed marketing tables, removed zero-width characters. It does not attempt visual fidelity to Outlook's web interface. Headers become YAML frontmatter. Quoted replies are truncated at the first recognized quote separator. Inline images become placeholder annotations rather than broken `<img>` tags.

This is a different design target from "convert email to readable document." The consumer is a grep pipeline, a vector index, or an LLM context window — not a person scrolling through their inbox.

### Constrained scope is a feature

The system deliberately limits itself. No attachment content extraction. No folder state synchronization. No multi-account management. No background daemon. No arbitrary Graph passthrough.

These are not missing features. Each boundary reduces the surface area that must be tested, secured, and maintained. A local, single-user, read-first CLI with explicit write commands has a dramatically smaller failure domain than a full email client or synchronization service.

### Default folders reflect real usage

The default download targets are `Inbox` and `Archive`. For most Outlook.com users, email does not live in a single folder — the archive is an active workspace, not cold storage. The default behavior matches this pattern. Explicit `--folder` flags override the defaults.

## What it does not do

- Modify message state (read/unread, move, delete) on the server
- Extract or render attachment content
- Build a local full-text search index
- Run as a background sync daemon or webhook receiver
- Manage multiple accounts or users
- Pass through arbitrary Graph API calls
- Full calendar synchronization, event updates, deletions, or RSVP management
- Mailbox settings management

## Module map

```
src/outlook_skill/
├── config.py         — env loading, Settings dataclass, defaults
├── auth.py           — MSAL public client, token cache, login/refresh/status
├── graph_client.py   — folder resolution, paginated listing, MIME download
├── downloader.py     — folder iteration, idempotency, header extraction, progress reporting
├── store.py          — SQLite schema, CRUD, SHA-256 hashing, schema migration
├── markdown_renderer.py — parallel .eml→.md, HTML cleaning, quote stripping
├── exporter.py       — YAML frontmatter export, filtering, slugified filenames
├── sender.py         — standalone sendMail, body-format conversion, dry-run
├── replier.py        — createReply/createReplyAll draft flow, attachment upload, recipient override
├── calendar.py       — invite creation, event listing, recurring filtering
├── spam_triage.py    — rule-based labeling, multi-label classification
├── mail.py           — (thin) local read and list via store + downloader
├── models.py         — dataclasses: DownloadedMessage, SpamRule, TriageLabel
├── errors.py         — typed exceptions: ConfigError, AuthRequiredError, GraphApiError
├── cli.py            — argparse shell, progress reporting, output formatting
docs/
├── prd.md            — this document
├── rfc.md            — architecture and design rationale
├── skill_outlook.md  — see skills/skill_outlook.md
├── working.md        — development log
├── test.md           — test strategy
scripts/
└── outlook           — convenience wrapper
tests/
```

## Verification

The test pyramid has three layers: unit tests for pure logic, mocked integration tests for module seams, and live integration tests for the full OAuth2→Graph→local disk pipeline — the last layer defaults to skip and requires explicit env-var gating. Write operations (send, invite) require additional allow flags beyond the base live test flag. Reply dry-runs may create drafts but do not send mail. A smoke check script exercises the CLI contract end-to-end.
