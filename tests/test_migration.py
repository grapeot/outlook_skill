import hashlib
import sqlite3
from pathlib import Path

import pytest

from outlook_skill.config import Settings
from outlook_skill.errors import GraphApiError, OutlookSkillError
from outlook_skill.migration import migrate_current_state
from outlook_skill.store import MailStore


OLD_MESSAGES_SCHEMA = """
CREATE TABLE messages (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    account TEXT NOT NULL,
    folder TEXT NOT NULL,
    uid INTEGER NOT NULL,
    uidvalidity INTEGER,
    message_id TEXT,
    subject TEXT,
    from_addr TEXT,
    received_at TEXT,
    internaldate TEXT,
    size INTEGER,
    flags_json TEXT NOT NULL,
    mime_path TEXT NOT NULL,
    downloaded_at TEXT NOT NULL,
    sha256 TEXT NOT NULL,
    UNIQUE(account, folder, uidvalidity, uid)
);
CREATE TABLE triage_labels (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    message_uid TEXT NOT NULL,
    folder TEXT NOT NULL,
    label TEXT NOT NULL,
    rule_id INTEGER,
    confidence REAL,
    labeled_at TEXT NOT NULL,
    UNIQUE(message_uid, folder, label)
);
"""


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


def insert_legacy(connection, row_id, folder, uid, raw_name, sha, mime_path):
    connection.execute(
        """
        INSERT INTO messages (
            id, account, folder, uid, uidvalidity, message_id, subject, from_addr,
            received_at, internaldate, size, flags_json, mime_path, downloaded_at, sha256
        ) VALUES (?, 'duck@example.com', ?, ?, 'graph', '<m@example.com>', 'Subject',
                  'sender@example.com', '2026-01-01T00:00:00Z', '2026-01-01T00:00:00Z',
                  10, '', ?, ?, ?)
        """,
        (row_id, folder, uid, str(mime_path), f"2026-01-0{row_id}T00:00:00Z", sha),
    )


def legacy_db(settings: Settings) -> None:
    settings.messages_dir.mkdir(parents=True)
    connection = sqlite3.connect(settings.db_path)
    connection.executescript(OLD_MESSAGES_SCHEMA)
    insert_legacy(connection, 1, "INBOX", "same-id", "old", "sha-old", settings.messages_dir / "old.eml")
    insert_legacy(connection, 2, "Archive", "same-id", "new", "sha-new", settings.messages_dir / "new.eml")
    insert_legacy(connection, 3, "INBOX", "unique-id", "unique", "sha-unique", settings.messages_dir / "unique.eml")
    connection.commit()
    connection.close()


class FakeGraph:
    def __init__(self, settings):
        pass

    def get_message_metadata(self, *, message_id):
        return {"id": message_id, "parentFolderId": "archive-folder-id"}

    def get_folder(self, *, folder_id):
        return {"id": folder_id, "displayName": "Archive"}

    def fetch_message_mime(self, *, message_id):
        return b"Subject: Current\nFrom: sender@example.com\n\nCurrent body"

    def close(self):
        pass


def test_migration_dry_run_has_no_graph_or_schema_mutation(monkeypatch, tmp_path):
    settings = make_settings(tmp_path)
    legacy_db(settings)
    monkeypatch.setattr(
        "outlook_skill.migration.GraphMailClient",
        lambda settings: (_ for _ in ()).throw(AssertionError("dry-run constructed Graph client")),
    )

    result = migrate_current_state(settings)

    assert result["source_rows"] == 3
    assert result["canonical_rows"] == 2
    assert result["duplicate_groups"] == 1
    assert result["duplicate_groups_with_different_content"] == 1
    columns = {row[1] for row in sqlite3.connect(settings.db_path).execute("PRAGMA table_info(messages)")}
    assert "parent_folder_id" not in columns


def test_migration_deduplicates_current_graph_id_and_preserves_current_content(monkeypatch, tmp_path):
    settings = make_settings(tmp_path)
    legacy_db(settings)
    monkeypatch.setattr("outlook_skill.migration.GraphMailClient", FakeGraph)

    result = migrate_current_state(settings, execute=True)

    assert result["migrated"] is True
    assert result["source_rows"] == 3
    assert result["canonical_rows"] == 2
    store = MailStore(settings)
    try:
        current = store.get_message(account="duck@example.com", uid="same-id")
        assert current["folder"] == "Archive"
        assert current["parent_folder_id"] == "archive-folder-id"
        assert Path(current["mime_path"]).read_bytes() == b"Subject: Current\nFrom: sender@example.com\n\nCurrent body"
        assert store.get_message(account="duck@example.com", uid="unique-id") is not None
        assert store.connection.execute("SELECT count(*) FROM messages").fetchone()[0] == 2
    finally:
        store.close()

    assert migrate_current_state(settings)["already_migrated"] is True


def test_migration_keeps_latest_local_candidate_for_unresolved_404(monkeypatch, tmp_path):
    settings = make_settings(tmp_path)
    legacy_db(settings)

    class MissingGraph(FakeGraph):
        def get_message_metadata(self, *, message_id):
            raise GraphApiError("missing", status_code=404)

    monkeypatch.setattr("outlook_skill.migration.GraphMailClient", MissingGraph)

    result = migrate_current_state(settings, execute=True)

    assert result["unresolved_duplicate_groups"] == 1
    store = MailStore(settings)
    try:
        current = store.get_message(account="duck@example.com", uid="same-id")
        assert current["folder"] == "Archive"
        assert current["parent_folder_id"] is None
        assert current["mime_path"].endswith("new.eml")
    finally:
        store.close()


def test_migration_mime_paths_include_message_identity(monkeypatch, tmp_path):
    settings = make_settings(tmp_path)
    legacy_db(settings)
    connection = sqlite3.connect(settings.db_path)
    insert_legacy(connection, 4, "INBOX", "other-id", "old-other", "sha-other-old", settings.messages_dir / "old-other.eml")
    insert_legacy(connection, 5, "Archive", "other-id", "new-other", "sha-other-new", settings.messages_dir / "new-other.eml")
    connection.commit()
    connection.close()
    monkeypatch.setattr("outlook_skill.migration.GraphMailClient", FakeGraph)

    migrate_current_state(settings, execute=True)

    store = MailStore(settings)
    try:
        first = store.get_message(account="duck@example.com", uid="same-id")
        second = store.get_message(account="duck@example.com", uid="other-id")
        assert first["mime_path"] != second["mime_path"]
    finally:
        store.close()


def test_migration_rejects_singleton_collision_without_replacing_legacy(monkeypatch, tmp_path):
    settings = make_settings(tmp_path)
    legacy_db(settings)

    class CollisionGraph(FakeGraph):
        def get_message_metadata(self, *, message_id):
            return {"id": "unique-id", "parentFolderId": "archive-folder-id"}

    monkeypatch.setattr("outlook_skill.migration.GraphMailClient", CollisionGraph)

    with pytest.raises(RuntimeError, match="collides with an existing singleton"):
        migrate_current_state(settings, execute=True)

    connection = sqlite3.connect(settings.db_path)
    columns = {row[1] for row in connection.execute("PRAGMA table_info(messages)")}
    assert "parent_folder_id" not in columns
    assert connection.execute("SELECT count(*) FROM messages").fetchone()[0] == 3
    connection.close()


def test_migration_dry_run_does_not_create_missing_database(tmp_path):
    settings = make_settings(tmp_path)

    with pytest.raises(sqlite3.OperationalError):
        migrate_current_state(settings)

    assert not settings.db_path.exists()


def test_mailstore_rejects_legacy_schema(tmp_path):
    settings = make_settings(tmp_path)
    legacy_db(settings)

    with pytest.raises(OutlookSkillError, match="migrate-current-state"):
        MailStore(settings)
