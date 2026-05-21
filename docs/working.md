# Working Notes

## Changelog

### 2026-05-20

- Added `mail reply-all` backed by Microsoft Graph `createReplyAll`, reusing the existing reply draft flow for body patching, recipient overrides, attachments, dry-run summaries, and final send.
- Added `operation` to reply JSON results so dry-run and sent output distinguishes `reply` from `reply_all`.
- Added library and CLI tests for reply-all endpoint selection, CLI dispatch, and dry-run not sending the draft.

### 2026-05-19

- Changed `calendar invite` to create attendee-less appointments by default. `--to` is now optional; when omitted the Graph payload uses an empty `attendees` list. Passing `--to` or `--optional-attendee` still includes attendees exactly as before. Updated library validation and tests accordingly.

### 2026-05-05

- Added standalone `mail send` using Microsoft Graph `POST /me/sendMail`, with `--dry-run` as the default safety gate.
- Added `calendar invite` using `POST /me/calendar/events` to create events with attendees; Graph sends invitations automatically.
- Added sender/calendar unit tests, CLI delegation tests, and a default-skipped live integration skeleton. Real email sending requires `OUTLOOK_LIVE_ALLOW_SEND=1`; real calendar invite creation requires `OUTLOOK_LIVE_ALLOW_CALENDAR_INVITE=1`.
- Updated scope, permissions, and non-goal boundaries across PRD, RFC, test, skill, and README. Calendar invite requires delegated `Calendars.ReadWrite` and re-running `auth login` for consent.

### 2026-04-14

- Created the repo and project skeleton (initial directory structure for a Graph-based mail skill).
- Initialized an independent git repo and project-level `.venv`.
- Built an initial Graph mail + calendar MVP and completed first-round tests and commits.
- Based on evolving requirements, converged the project to an IMAP-only mail download tool.
- Rewrote `README.md`, `AGENTS.md`, `docs/prd.md`, `docs/rfc.md`, `docs/test.md`, removing all Graph and calendar assumptions.
- Removed `auth.py`, `client.py`, `calendar.py`, and corresponding calendar tests.
- Added IMAP versions of `config.py`, `models.py`, `imap_client.py`, `downloader.py`, `store.py`, `mail.py`, `cli.py`.
- Converged the CLI to three commands: `doctor config`, `mail download`, `mail list-local`.
- Converged local persistence to raw `.eml` + SQLite metadata.
- After confirming that `LOGIN` and app password approaches fail on real Outlook.com accounts, switched authentication to OAuth2/XOAUTH2.
- Restored `msal` dependency, added `auth login` and `auth status`, connected IMAP via `AUTHENTICATE XOAUTH2`.
- Unified `.env.example`, README, PRD, RFC, and test strategy around the app registration + token cache model.
- Switched the transport layer from IMAP to Microsoft Graph for mail download.
- Added `graph_client.py`, replacing IMAP fetches with `Mail.Read` and the MIME `/$value` endpoint.
- Kept the local `.eml` + SQLite storage layer, switching only the idempotency key semantics to Graph message ID.
- Following Oracle review, documented that Phase 1 relies on Graph message ID for idempotency; long-term sync stability would need separate immutable ID evaluation.
- Completed first Graph authorization on a real Outlook.com account and verified successful download of the most recent week of mail.
- Added download progress bars: JSON mode outputs progress to stderr, keeping stdout machine-parseable.
- Extended default download scope from Inbox-only to Inbox + Archive, with `--folder` override support.
- Positioned the CLI as the stable execution backend for the workspace skill; added skill documentation and skill index registration.
- Added `markdown_renderer.py` and the `mail render-markdown` CLI, with parallel `.eml` → AI-oriented Markdown rendering.
- Completed first-round sample rendering of 50 messages, distributed across 5 reviewers for batch review.
- Based on first-round review: added zero-width character cleaning, tracking URL simplification, marketing layout table unwrapping, and more robust batch error handling.
- Completed second-round sample rendering of another 50 messages, distributed again across 5 reviewers.
- Added `spam_triage.py`, `spam_rules`, `triage_labels`, and the `triage` CLI to validate a rule-based email cleaning workflow.
- Two rounds of high-confidence `spam` rules covered 7,404 / 38,290 messages, approximately 19.34%.
- A third round of `low_value` rules added 3,863 more, bringing combined `spam + low_value` coverage to 11,267 / 38,290, approximately 29.43%.

## Lessons Learned

- For a local archiving and indexing-preparation use case, IMAP is a more natural fit than Graph in principle — but Graph's delegated permissions and consistent API surface proved more practical for personal Outlook.com accounts.
- Raw MIME is the core asset. It takes priority over full-text indexing and attachment parsing because every downstream transformation can be rebuilt from raw, but the reverse is impossible.
- Outlook.com's IMAP documentation points toward OAuth2/Modern Auth, but a working download loop needs to ship before chasing the full OAuth2 path. Graph sidesteps this by making auth a first-class part of the API.
- For a once-daily downloader, pulling a recent time window and using `(folder, uidvalidity, uid)` for idempotency is sufficient. Full delta sync is overkill for this use case.
- On personal Outlook.com accounts, XOAUTH2 requires the same app registration prerequisites as Graph but with a more fragile token exchange path.
- For personal Outlook.com, Graph's delegated `Mail.Read` provides a cleaner permission model than IMAP and maps directly to the download-then-render pipeline.
- The biggest long-term caveat with Graph is not authentication but message identifier stability — Graph message IDs can change in mobile sync scenarios, which matters for long-term synchronization across folder moves.
- For CLI tools, progress feedback must be separated from machine-readable output. JSON mode must never let progress bars pollute stdout.
- For email archiving, users intuitively think of their mailbox as Inbox + Archive, not just Inbox. The default folder strategy must match this mental model.
- First-round sample review identified table layouts, tracking URLs, zero-width characters, footer disclaimers, and HTML-only marketing emails as the primary sources of Markdown rendering noise.
- Second-round review confirmed that noise filtering had significantly improved. Remaining issues centered on marketing email card/table structure preservation, further tracking URL compression, and inline image placeholder quality.
- Rule-based triage converges quickly in early rounds but hits diminishing returns as the problem shifts from obvious spam to gray-zone low-value notifications. Multi-label classification (not binary spam/not-spam) is the right model for this transition.
