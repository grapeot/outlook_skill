# Outlook Skill

A local CLI tool that lets AI agents download, read, search, send, and triage Outlook.com email. It authenticates via OAuth2, talks to Microsoft Graph, and stores everything locally as `.eml` + SQLite + Markdown.

This is **not an email client**. It is not built for humans to read their inbox. It is built so that an AI agent — running in your terminal, via Claude Code or Cursor or any other agent framework — can handle your Outlook email without you opening a browser.

## Why this works for personal accounts

Every Outlook-to-AI connector needs a Microsoft Entra app registration behind it. Entra is only available to Office 365 enterprise subscribers — which means the vast majority of personal Outlook.com / Hotmail / Live users cannot create their own app registration. Without it, the OAuth2 handshake fails before it starts.

Most connectors stop here. They work for work/school accounts and silently fail for personal accounts.

This project is different: it ships with a pre-deployed Entra app that has "personal Microsoft accounts" enabled. You do not need Entra access. You do not need an Office 365 subscription. You copy the `.env.example`, and the client ID is already there.

There is a tradeoff. The client ID is shared — everyone using this project authenticates through the same app registration. This is the standard model for open-source desktop apps (MSAL public client, no client secret, interactive browser login). Your token never touches our servers and is stored only on your local machine. But:

- **Rate limits** are aggregated across all users of the shared client ID. If someone abuses it, Microsoft may throttle the app, affecting everyone.
- **The consent screen** shows the app name we registered, not yours. Users see our publisher name when they authorize.

If these matter to you, **bring your own client ID**. You need a Microsoft Entra app registration (work/school account required to access the Entra portal). Create an app with personal account support, `http://localhost` redirect URI, and `Mail.Read` / `Mail.ReadWrite` / `Mail.Send` / `Calendars.ReadWrite` delegated permissions. Replace the `OUTLOOK_CLIENT_ID` in `.env` with yours. If you don't have Entra access, use the provided client ID — it works.

## Install (for the human)

Three steps. Do them once.

**1. Install the package.**

```bash
git clone <this-repo> outlook_skill
cd outlook_skill
cp .env.example .env
# Edit .env: fill in OUTLOOK_EMAIL
# OUTLOOK_CLIENT_ID is pre-filled — replace it if you're using your own
uv venv .venv
uv pip install --python .venv/bin/python -e '.[dev]'
```

**2. Copy the skill file to where your AI agent reads skills.**

The file is `skills/skill_outlook.md`. Copy it anywhere your agent framework looks for skill files — alongside your other `.md` skill files, in a `rules/` or `skills/` directory, or wherever your agent's config expects.

**3. Tell your AI agent about it.**

Add a line to your `AGENTS.md`, `CLAUDE.md`, or equivalent agent config file:

> If the user asks about Outlook email (downloading, reading, searching, sending, calendar), read `skills/skill_outlook.md` first, then follow its instructions.

That's it. The AI handles everything else: first-time login (`auth login` opens a browser), downloading mail, rendering markdown, send/reply/reply-all/calendar operations.

## For AI Agents

When a user asks you to do anything with Outlook email, open `skills/skill_outlook.md` in this repo (or wherever the human copied it). That file contains the full CLI contract: every command, every flag, every output format, the standard workflow, and known caveats. Do not guess commands — read the skill file.

What you need to know up front:

- **Entry point**: `.venv/bin/python -m outlook_skill.cli <command>` (from the repo root), or `scripts/outlook <command>`
- **JSON output**: add `--format json` to any command for machine-readable output. Progress bars go to stderr; results go to stdout.
- **First-time setup**: the human needs to run `auth login` once (opens a browser). After that, token refresh is automatic.
- **Standard pipeline**: `mail download` → `mail export-md` → grep/search the markdown directory → `mail read --graph-id <id>` for full body → `mail reply`, `mail reply-all`, or `mail send` to compose responses.
- **Safety**: all write commands (`send`, `reply`, `reply-all`, `invite`) support `--dry-run`. Use it before executing.

When the human hasn't completed setup, tell them clearly what's missing. Common issues:
- No `.env`: ask them to copy `.env.example` and edit `OUTLOOK_EMAIL`
- No token cache: tell them to run `auth login`
- Wrong permissions: tell them which scope is missing (`Mail.Read`, `Mail.Send`, etc.)
- Using the shared client ID and hitting rate limits: suggest they deploy their own app registration if they have Entra access

## For Developers

If you want to modify the library or CLI:

```bash
# Run tests (no network)
.venv/bin/python -m pytest -v

# Run live integration tests (requires valid OAuth2 config)
OUTLOOK_ENABLE_LIVE_TESTS=1 .venv/bin/python -m pytest -v -m live_integration

# Smoke check after changes
.venv/bin/python -m outlook_skill.cli doctor config --format json
.venv/bin/python -m outlook_skill.cli mail list-local --limit 5 --format json
```

Write operations in tests require additional flags: `OUTLOOK_LIVE_ALLOW_SEND=1` and `OUTLOOK_LIVE_ALLOW_CALENDAR_INVITE=1`. No test sends real email without explicit opt-in.

The library lives in `src/outlook_skill/` (10 modules). The CLI is a thin argparse shell. Architecture and design rationale: [docs/rfc.md](docs/rfc.md). Product description: [docs/prd.md](docs/prd.md).

### Local data layout

```
data/mail/
├── messages/                # Raw .eml per message
├── markdowns/               # AI-oriented .md (parallel renderer output)
├── markdown/                # YAML frontmatter .md (exporter output)
├── mail.db                  # SQLite: messages, spam_rules, triage_labels
└── oauth_token_cache.json   # MSAL token cache with refresh token
```

All data stays local. Nothing is uploaded or synced to any server beyond the Graph API calls to Microsoft.
