from pathlib import Path

from outlook_skill.config import Settings
from outlook_skill.store import MailStore


def make_settings(tmp_path: Path) -> Settings:
    data_dir = tmp_path / "mail"
    return Settings(
        email="duck@example.com",
        client_id="client-id",
        authority="https://login.microsoftonline.com/consumers",
        scopes=("Mail.Read", "offline_access"),
        graph_base_url="https://graph.microsoft.com/v1.0",
        default_folders=("INBOX", "Archive"),
        token_cache_path=data_dir / "oauth_token_cache.json",
        data_dir=data_dir,
        messages_dir=data_dir / "messages",
        markdowns_dir=data_dir / "markdowns",
        db_path=data_dir / "mail.db",
    )


def test_store_saves_message_and_detects_existing(tmp_path):
    settings = make_settings(tmp_path)
    store = MailStore(settings)
    try:
        saved = store.save_message(
            account="duck@example.com",
            folder="INBOX",
            uid="graph-message-id",
            uidvalidity="graph",
            raw_message=b"Subject: Hello\n\nBody",
            message_id="<abc@example.com>",
            subject="Hello",
            from_addr="duck@example.com",
            received_at="Tue, 14 Apr 2026 10:00:00 +0000",
            internaldate="14-Apr-2026 10:00:00 +0000",
            size=20,
            flags_json="\\Seen",
        )

        assert Path(saved.mime_path).exists()
        assert store.has_message(account="duck@example.com", folder="INBOX", uidvalidity="graph", uid="graph-message-id") is True
    finally:
        store.close()
