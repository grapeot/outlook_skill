import hashlib
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
        assert store.has_message(account="duck@example.com", uid="graph-message-id") is True
    finally:
        store.close()


def _save(store: MailStore, *, folder: str, uid: str, raw: bytes = b"Subject: Hello\n\nBody") -> None:
    store.save_message(
        account="duck@example.com",
        folder=folder,
        uid=uid,
        uidvalidity="graph",
        raw_message=raw,
        message_id="<abc@example.com>",
        subject="Hello",
        from_addr="duck@example.com",
        received_at="Tue, 14 Apr 2026 10:00:00 +0000",
        internaldate="14-Apr-2026 10:00:00 +0000",
        size=len(raw),
        flags_json="",
    )


def test_same_graph_id_updates_current_folder_without_duplicate(tmp_path):
    store = MailStore(make_settings(tmp_path))
    try:
        _save(store, folder="INBOX", uid="graph-message-id")
        first_path = store.get_message(account="duck@example.com", uid="graph-message-id")["mime_path"]

        _save(store, folder="Archive", uid="graph-message-id", raw=b"Subject: Changed\n\nBody")
        matches = store.find_messages(graph_id="graph-message-id")

        assert len(matches) == 1
        assert matches[0]["folder"] == "Archive"
        assert matches[0]["mime_path"] == first_path
        assert matches[0]["sha256"] == hashlib.sha256(b"Subject: Changed\n\nBody").hexdigest()
    finally:
        store.close()


def test_update_message_location_tracks_returned_graph_id(tmp_path):
    store = MailStore(make_settings(tmp_path))
    try:
        _save(store, folder="INBOX", uid="old-id")
        original_raw = b"Subject: Hello\n\nBody"

        result = store.update_message_location(
            account="duck@example.com",
            old_graph_id="old-id",
            new_graph_id="new-id",
            folder="Archive",
            parent_folder_id="archive-folder-id",
        )

        assert store.get_message(account="duck@example.com", uid="old-id") is None
        current = store.get_message(account="duck@example.com", uid="new-id")
        assert current["folder"] == "Archive"
        assert current["parent_folder_id"] == "archive-folder-id"
        assert Path(result["mime_path"]).read_bytes() == original_raw
    finally:
        store.close()


def test_update_message_location_moves_triage_labels_with_message(tmp_path):
    store = MailStore(make_settings(tmp_path))
    try:
        _save(store, folder="INBOX", uid="old-id")
        store.save_triage_label(
            message_uid="old-id",
            folder="INBOX",
            label="spam",
            rule_id=None,
            confidence=0.9,
        )

        store.update_message_location(
            account="duck@example.com",
            old_graph_id="old-id",
            new_graph_id="new-id",
            folder="Archive",
            parent_folder_id="archive-folder-id",
        )

        labels = store.list_triage_labels()
        assert len(labels) == 1
        assert labels[0].message_uid == "new-id"
        assert labels[0].folder == "Archive"
    finally:
        store.close()
