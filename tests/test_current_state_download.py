from pathlib import Path

from outlook_skill.config import Settings
from outlook_skill.downloader import download_recent_mail
from outlook_skill.store import MailStore


def make_settings(tmp_path: Path) -> Settings:
    data_dir = tmp_path / "mail"
    return Settings(
        email="duck@example.com",
        client_id="client-id",
        authority="https://login.microsoftonline.com/consumers",
        scopes=("Mail.ReadWrite", "offline_access"),
        graph_base_url="https://graph.microsoft.com/v1.0",
        default_folders=("INBOX", "Archive"),
        token_cache_path=data_dir / "oauth_token_cache.json",
        data_dir=data_dir,
        messages_dir=data_dir / "messages",
        markdowns_dir=data_dir / "markdowns",
        db_path=data_dir / "mail.db",
    )


class FakeGraph:
    def __init__(self, settings):
        pass

    def resolve_folder(self, name):
        return ("archive-folder-id", "Archive")

    def list_recent_messages(self, *, folder_id, days, limit):
        return [{"id": "same-id", "received_at": "2026-01-01T00:00:00Z", "parent_folder_id": folder_id}]

    def fetch_message_mime(self, *, message_id):
        return b"Subject: Current\n\nBody"

    def close(self):
        pass


def test_download_refreshes_folder_instead_of_inserting_duplicate(monkeypatch, tmp_path):
    settings = make_settings(tmp_path)
    store = MailStore(settings)
    try:
        store.save_message(
            account=settings.email,
            folder="INBOX",
            uid="same-id",
            uidvalidity="graph",
            raw_message=b"Subject: Current\n\nBody",
            message_id="<m@example.com>",
            subject="Current",
            from_addr="sender@example.com",
            received_at="2026-01-01T00:00:00Z",
            internaldate="2026-01-01T00:00:00Z",
            size=10,
            flags_json="",
        )
    finally:
        store.close()
    monkeypatch.setattr("outlook_skill.downloader.GraphMailClient", FakeGraph)

    result = download_recent_mail(settings, days=1, limit=10, folders=("Archive",))

    assert result["updated_count"] == 1
    store = MailStore(settings)
    try:
        current = store.get_message(account=settings.email, uid="same-id")
        assert current["folder"] == "Archive"
        assert current["parent_folder_id"] == "archive-folder-id"
        assert store.connection.execute("SELECT count(*) FROM messages").fetchone()[0] == 1
    finally:
        store.close()
