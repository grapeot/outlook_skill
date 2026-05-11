# Outlook Skill — Architecture and Design Rationale

## Scope

The system covers a deliberately narrow set of Microsoft Graph operations on personal Outlook.com accounts: OAuth2 authentication, mail folder listing, MIME download, local `.eml` + SQLite storage, `.eml` → Markdown rendering, standalone email sending, in-thread reply, and calendar invite creation with event listing.

This scope serves two goals. First, it gives AI workflows a stable, reusable entry point to a personal Outlook.com mailbox — the commands an agent needs to download, read, search, and optionally send email. Second, it draws a hard line at the skill layer: the system never grows into a full mail client or calendar application.

## Module boundaries

The codebase is organized into ten library modules plus a thin CLI. Each boundary reflects a real architectural seam — a place where responsibilities change, dependencies invert, or test strategies diverge.

### 1. `config.py` — Configuration resolution

Loads `.env` from the project root, reads environment variables, resolves relative paths against `PROJECT_ROOT`, and assembles a frozen `Settings` dataclass. This module is the single choke point for all runtime configuration. Defaults are defined as module-level constants: `DEFAULT_FOLDERS = ("INBOX", "Archive")`, `DEFAULT_AUTHORITY = "https://login.microsoftonline.com/consumers"`, `DEFAULT_DATA_DIR = "data/mail"`.

The `doctor_info()` function provides a diagnostic view of the configuration state without requiring network access or valid credentials. It reports masked values for sensitive fields and checks file existence for token cache, data directory, and database paths.

### 2. `auth.py` — OAuth2 token management

Wraps MSAL's `PublicClientApplication` with a serializable token cache. On construction, it reads the cache from disk if present. `login()` supports two flows: interactive browser (`acquire_token_interactive` with `select_account` prompt and port 0 for automatic redirect URI) and device code (`initiate_device_flow` + `acquire_token_by_device_flow`). Both persist the cache to disk on success.

`get_access_token()` implements silent refresh: it looks up cached accounts by email, falls back to any cached account, calls `acquire_token_silent`, and raises `AuthRequiredError` if the token is missing or expired. The 401 retry in `graph_client.py` triggers a client rebuild that calls `get_access_token()` again, effectively refreshing the token mid-session.

`status()` reports whether the token cache contains accounts matching the configured email, which is the precondition for silent token acquisition to succeed.

### 3. `graph_client.py` — Microsoft Graph transport

A stateful HTTP client wrapping `httpx.Client` with Graph-specific behavior. It resolves folders by well-known name (`inbox`, `archive`, `sentitems`, `drafts`, `deleteditems`) or by `displayName` filter with single-quote escaping. Message listing uses `$filter=receivedDateTime ge {cutoff}` with `$orderby=receivedDateTime desc` and OData pagination (`@odata.nextLink`). MIME download hits `/$value` with the ImmutableId header.

The key design choice is the 401 retry pattern: when any request returns 401, the client calls `_refresh_client()`, which rebuilds the `httpx.Client` with a fresh token from `AuthManager.get_access_token()`, and retries the request once. This means callers never need to handle token expiration — the transport layer absorbs it.

### 4. `downloader.py` — Download orchestration

The central pipeline that connects Graph transport to local storage. For each target folder, it resolves the folder ID, lists recent messages, checks idempotency via `store.has_message()`, fetches MIME for new messages, extracts headers via `email.parser.BytesParser`, and writes both `.eml` and SQLite record.

The progress callback protocol uses keyword-only arguments (`current`, `total`, `downloaded`, `skipped`) and separates progress rendering from download logic. The CLI instantiates a `tqdm`-based callback; library callers can provide their own.

### 5. `store.py` — Local persistence

SQLite with three tables defined in a single `SCHEMA` string. The `messages` table uses `(account, folder, uidvalidity, uid)` as a composite unique constraint — the same key used for idempotency. The `uid` column is typed as `INTEGER` but stores Graph message ID strings; this is a legacy artifact from the earlier IMAP uid schema and is handled at the application layer.

Schema migration runs on every `MailStore` instantiation via `_migrate_schema()`, which checks for missing columns (currently `label` on `spam_rules`) and adds them with `ALTER TABLE`. This is intentionally minimal — the project does not use a migration framework because the schema surface is small and changes are infrequent.

File paths use sanitized components: `@` → `_at_`, `/` → `_`, `=` → `_`. The filename includes account, folder, first 24 chars of the Graph ID, and first 12 chars of the SHA-256 hash.

### 6. `markdown_renderer.py` — AI-oriented rendering

Parallel `.eml` → `.md` conversion using `concurrent.futures.ThreadPoolExecutor`. Each worker parses the `.eml` with `email.parser.BytesParser`, selects the best body part (plain text preferred, HTML as fallback), converts HTML to Markdown via `markdownify`, and applies a pipeline of cleaning transformations:

- **Quote stripping**: terminates body text at recognized quote separators (`On ... wrote:`, `From:/Sent:/To:/Subject:` headers, `--- Original Message ---`)
- **Zero-width character removal**: regex strips U+200B, U+200C, U+200D, U+FEFF, U+2060
- **Marketing layout unwrapping**: `<table>` elements with >8 rows or >45% empty cells are collapsed to bullet lists
- **Tracking URL simplification**: URLs containing known tracking hosts (Mailgun, SendGrid, Klaviyo, etc.) are resolved to their redirect target via query parameter extraction (`url=`, `u=`, `target=`, `redirect=`, `redirect_url=`, `destination=`)
- **Hidden element removal**: `display:none` elements, `<script>`, `<style>`, `<head>`, `<noscript>`, `<svg>` are decomposed
- **Inline image handling**: `cid:` images become annotated placeholders; tracking-pixel images are removed; regular images become `【图片：alt】` markers
- **Noise cleanup**: Markdown table separators, very long tracking links (>140 chars), and double-link patterns are stripped

The rendering is not lossless. It intentionally discards visual structure (CSS, layout tables, tracking infrastructure) and preserves semantic content (headers, body text, attachment metadata). The result is optimized for grep, vector embedding, and LLM context windows — not for human reading.

### 7. `exporter.py` — Structured Markdown export

Reads from SQLite metadata + `.eml` files, applies optional filters (days, folders, subject, from), and writes YAML frontmatter Markdown files. Filenames follow the pattern `{date}_{folder}_{subject-slug}_{short-graph-id}.md`. Existing files are skipped unless `--force` is specified.

The YAML frontmatter includes `graph_id`, `folder`, `from`, `to`, `cc`, `subject`, `date`, and `body_source` — enough metadata for an AI agent to correlate the local file back to the Graph origin.

### 8. `sender.py` — Standalone email sending

Constructs a Graph `sendMail` payload with `message` (subject, body with contentType, toRecipients, ccRecipients, bccRecipients), inline attachments ≤3MB using `fileAttachment` with base64-encoded `contentBytes`, and `saveToSentItems` control. Body format conversion handles Markdown→HTML via `markdown.markdown()` when `body_format` is `markdown` or `md`.

Dry-run mode (`MailSendAllowFlags.send`) constructs the full payload, logs a summary to stderr, and returns without calling Graph.

### 9. `replier.py` — In-thread reply

Uses Graph's `createReply` endpoint to build a reply draft from the parent message, then attaches files and sends. Attachments ≤3MB use inline `fileAttachment`; larger attachments use `createUploadSession` for chunked upload. Recipient overrides (`--to`, `--cc`) modify the draft's recipient list before sending.

### 10. `calendar.py` — Calendar operations

Two operations on a single endpoint. `create_calendar_invite()` posts to `/me/calendar/events` with `subject`, `start`/`end` (with timeZone), `location`, `body` (with contentType), and `attendees` (required + optional). `list_calendar_events()` queries `/me/calendar/events` with `$filter=start/dateTime ge {start}` (using the complex type path syntax), `$orderby=start/dateTime`, `$select` for relevant fields, and `$top=100` as a safety ceiling. Client-side filtering removes daily/weekly recurring events when `--skip-recurring` is specified.

### 11. `cli.py` — Thin CLI shell

An argparse-based command tree that parses arguments, calls library functions, and formats output. The CLI owns two responsibilities that the library does not: progress rendering and output formatting.

Progress bars use `tqdm` and write to stderr in JSON mode to keep stdout clean for machine parsing. The `emit_output()` function dispatches between JSON (pretty-printed, sorted keys) and text (key-value pairs, with nested structures serialized as JSON strings).

## Architecture decisions

### Why library-first

A monolithic `sync_outlook.py` script works for a single use case but fails under composition. When an AI agent needs to download mail, render Markdown, and search locally, it should call three library functions — not parse stdout from a shell command. When a test needs to verify auth behavior without touching the network, it should instantiate `AuthManager` with controlled inputs — not mock `subprocess.run`.

Library-first means the contract is a Python API, not a CLI output format. The CLI is a consumer, not the implementation. This distinction is not academic: it determines whether tests can isolate layers and whether future consumers (a web UI, a cron job, a different agent framework) can reuse the same logic.

### Why Microsoft Graph, not IMAP

The project started on IMAP and switched to Graph for practical reasons, not architectural preference.

Outlook.com personal accounts support IMAP, but the experience degrades quickly. App passwords are deprecated. XOAUTH2 requires the same Entra app registration as Graph but with a more fragile token exchange path. IMAP folder listing, UIDVALIDITY semantics, and MIME retrieval vary across server implementations.

Microsoft Graph, by contrast, provides a consistent REST API: `GET /me/mailFolders` for folder resolution, `GET /me/messages` with OData filtering for listing, `GET /me/messages/{id}/$value` for raw MIME. The delegated `Mail.Read` permission maps cleanly to personal Outlook.com accounts. The API surface is narrower but more predictable.

The tradeoff is that Graph message IDs are less stable than IMAP UIDs in certain edge cases — particularly when messages move between folders or when the mailbox is accessed from mobile clients that use a different sync protocol. For the current use case (downloading recent mail within a time window), this is acceptable. For long-term synchronization across folder moves, an immutable ID strategy would need separate evaluation.

### Why consumers authority, not organizations

Personal Outlook.com accounts authenticate against `https://login.microsoftonline.com/consumers`, not the `organizations` or tenant-specific endpoints. This is a deliberate constraint: the system targets individual users with personal Microsoft accounts, not enterprise tenants.

The distinction matters because many Entra app registration guides assume a work/school context. Selecting "Accounts in any organizational directory and personal Microsoft accounts" during app registration is necessary for personal accounts to complete consent. The `consumers` authority endpoint is not a fallback — it is the correct endpoint for this user population.

### Why raw MIME first

The download pipeline writes `.eml` to disk before extracting metadata or rendering Markdown. This ordering is not about performance. It is about data preservation.

Raw MIME contains the complete email: headers, body parts in all formats (text/plain, text/html), attachments (base64-encoded), and encoding metadata. Any downstream transformation — body extraction, Markdown rendering, attachment parsing — can be re-run from the raw `.eml`. The reverse is not true: you cannot reconstruct MIME headers from Markdown, and you cannot recover an attachment that was never saved.

This principle applies to the idempotency model as well. The system checks for existence before download: if `(account, folder, uidvalidity, uid)` already has a record in SQLite, the message is skipped. The `.eml` is the durable artifact; SQLite is the index.

### Why the Markdown renderer is aggressive

The renderer makes irreversible choices: it strips quotes, removes tracking infrastructure, collapses layout tables, and replaces images with text annotations. These choices are correct for the target consumer (AI agents doing search and summarization) but inappropriate for a general-purpose email-to-text converter.

The design was validated through two rounds of sampled rendering with human review. The first round surfaced the main noise sources: zero-width characters in marketing emails, tracking URLs spanning hundreds of characters, HTML-only messages with useless text/plain stubs, and layout tables that produce unreadable Markdown. The second round confirmed that the cleaning pipeline handles these cases without introducing new distortion on clean emails.

### Why the triage system uses rules, not ML

The spam triage system applies deterministic rules (`sender_exact`, `sender_domain`, `subject_regex`, `body_regex`) and persists labels to SQLite. It does not use ML classification, Bayesian scoring, or heuristic weighting.

The rationale is explainability and control. Each label is traceable to a specific rule. Rules can be added, removed, and audited without retraining. The workflow — identify a high-frequency sender family, write a narrow rule, test on a sample, apply to the full corpus — is manual but predictable. This matters when the goal is not maximum recall but maximum trust: the user knows exactly why each message was labeled.

The multi-label design (`spam`, `low_value`, etc.) reflects the observation that binary spam/not-spam classification hits diminishing returns quickly. The first few rules catch obvious spam (promotional senders, newsletters that were never subscribed to). After that, the harder problem is low-value notifications — receipts, shipping updates, calendar reminders — that are legitimate but not worth reading. Multi-label classification separates these categories rather than forcing them into a single spam bucket.

## Local data model

Four layers of local data:

| Layer | Format | Location | Purpose |
|---|---|---|---|
| Token cache | JSON | `data/mail/oauth_token_cache.json` | MSAL serialized cache with refresh token |
| Raw MIME | `.eml` files | `data/mail/messages/*.eml` | Canonical email copy |
| Metadata | SQLite | `data/mail/mail.db` | Indexed queryable fields |
| AI Markdown | `.md` files (YAML frontmatter) | `data/mail/markdown/` | AI-consumable text export |

The idempotency key is `(account, folder, uidvalidity, uid)` where `uidvalidity` is the string `"graph"` and `uid` is the Graph message ID. This schema preserves the column structure from the earlier IMAP implementation (`uidvalidity` was originally an IMAP concept) while adapting to Graph's identifier model.

Same-message-different-folder is treated as two separate records. If a message exists in both Inbox and Archive, it will be downloaded and stored twice under different folder keys. This is the correct behavior for the current use case (recent-time-window download) but would need revisiting for full-mailbox synchronization.

## CLI design

### Progress and output separation

When `--format json` is specified, the final result writes to stdout and progress bars write to stderr. This separation is not cosmetic — it ensures that shell pipelines (`| jq`, `> result.json`) receive valid JSON without progress bar noise. The same rule applies to both `mail download` and `mail render-markdown`.

### Dry-run as default safety

All write commands (`mail send`, `mail reply`, `calendar invite`) support `--dry-run`. In dry-run mode, the command constructs the full payload, validates inputs, and prints a summary — without calling the Graph write endpoint. This is the safety default for AI agents that should inspect before executing.

### Write gating in tests

Live integration tests default to skip. They require `OUTLOOK_ENABLE_LIVE_TESTS=1` to run. Write operations require additional allow flags: `OUTLOOK_LIVE_ALLOW_SEND=1` for email sending, `OUTLOOK_LIVE_ALLOW_CALENDAR_INVITE=1` for calendar invites. No test sends real email or creates real calendar events without explicit, separate opt-in.

## Error handling

The system defines typed exceptions in `errors.py`: `ConfigError` for missing environment variables, `AuthRequiredError` for missing or expired tokens, `GraphApiError` for HTTP errors with status code and response text, and `OutlookSkillError` as the general base.

The CLI catches `OutlookSkillError` at the top level, prints the message to stderr, and returns exit code 1. Library callers are expected to catch specific exception types for programmatic error handling.

Known failure modes that the system handles explicitly:

- Missing `OUTLOOK_EMAIL` or `OUTLOOK_CLIENT_ID`: `ConfigError` with specific message
- Expired or missing token: `AuthRequiredError` directing user to run `auth login`
- Graph 401 during a request: automatic token refresh and single retry in `graph_client.py`
- Folder not found: `GraphApiError` with the folder name in the message
- Email without plain text or HTML body: fallback text "(no plain text or html body available)"

## Deferred

Capabilities explicitly excluded from current scope, with rationale:

- **Attachment content extraction**: requires format-specific parsers (PDF, image, Office documents) that would expand the dependency surface and testing burden
- **Full-text search index**: SQLite `LIKE` queries are sufficient for the current message volume; a dedicated search engine would add infrastructure complexity without proportional benefit
- **Server-side message state modification** (read/unread, move, delete): introduces risk of unintended state changes when used by AI agents; the read-only default is a safety property
- **Background sync daemon**: adds process lifecycle management, scheduling, and failure recovery that belong in infrastructure outside this library
- **Multi-account management**: each instance serves one account; multi-account orchestration belongs in the calling infrastructure
- **Calendar full sync, event update, delete, RSVP**: each of these involves state management and notification semantics that differ from the current read + create model
