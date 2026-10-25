from pathlib import Path

import httpx

from outlook_skill.config import Settings
from outlook_skill.mover import move_message


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


class FakeStore:
    def __init__(self, cached=True):
        self.cached = cached
        self.updates = []

    def get_message(self, *, account, uid):
        return {"uid": uid} if self.cached else None

    def update_message_location(self, **kwargs):
        self.updates.append(kwargs)
        return kwargs

    def close(self):
        pass


class FakeGraph:
    def __init__(self, settings, *, returned_id="new-id", actual_parent="archive-id"):
        self.returned_id = returned_id
        self.actual_parent = actual_parent
        self.moves = []

    def resolve_folder(self, destination):
        return ("archive-id", "Archive") if destination == "archive" else ("inbox-id", "Inbox")

    def get_message_metadata(self, *, message_id):
        return {"id": message_id, "parentFolderId": self.actual_parent}

    def move_message(self, *, message_id, destination_folder_id):
        self.moves.append((message_id, destination_folder_id))
        if self.actual_parent != "unexpected":
            self.actual_parent = destination_folder_id
        return {"id": self.returned_id}

    def close(self):
        pass


def test_move_plan_constructs_no_clients(monkeypatch, tmp_path):
    monkeypatch.setattr("outlook_skill.mover.MailStore", lambda settings: (_ for _ in ()).throw(AssertionError()))
    monkeypatch.setattr("outlook_skill.mover.GraphMailClient", lambda settings: (_ for _ in ()).throw(AssertionError()))

    result = move_message(make_settings(tmp_path), graph_id="old-id", destination="archive")

    assert result["execute"] is False
    assert result["server_moved"] is False
    assert result["local_state"] == "not_run"


def test_move_uses_returned_id_and_reconciles_one_cached_row(monkeypatch, tmp_path):
    store = FakeStore(cached=True)
    graph = FakeGraph(None)
    monkeypatch.setattr("outlook_skill.mover.MailStore", lambda settings: store)
    monkeypatch.setattr("outlook_skill.mover.GraphMailClient", lambda settings: graph)

    result = move_message(make_settings(tmp_path), graph_id="old-id", destination="archive", execute=True)

    assert graph.moves == [("old-id", "archive-id")]
    assert result["old_graph_id"] == "old-id"
    assert result["new_graph_id"] == "new-id"
    assert result["readback_verified"] is True
    assert result["local_state"] == "reconciled"
    assert result["completed"] is True
    assert store.updates[0]["old_graph_id"] == "old-id"
    assert store.updates[0]["new_graph_id"] == "new-id"
    assert store.updates[0]["parent_folder_id"] == "archive-id"


def test_move_readback_mismatch_does_not_update_local(monkeypatch, tmp_path):
    store = FakeStore(cached=True)
    graph = FakeGraph(None, actual_parent="unexpected")
    monkeypatch.setattr("outlook_skill.mover.MailStore", lambda settings: store)
    monkeypatch.setattr("outlook_skill.mover.GraphMailClient", lambda settings: graph)

    result = move_message(make_settings(tmp_path), graph_id="old-id", destination="archive", execute=True)

    assert result["server_moved"] is True
    assert result["readback_verified"] is False
    assert result["local_state"] == "pending"
    assert result["completed"] is False
    assert store.updates == []


def test_move_readback_timeout_returns_current_id_and_pending(monkeypatch, tmp_path):
    store = FakeStore(cached=True)

    class TimeoutGraph(FakeGraph):
        def get_message_metadata(self, *, message_id):
            if message_id == "new-id":
                raise httpx.ReadTimeout("readback timed out")
            return {"id": message_id, "parentFolderId": "archive-id"}

    graph = TimeoutGraph(None)
    monkeypatch.setattr("outlook_skill.mover.MailStore", lambda settings: store)
    monkeypatch.setattr("outlook_skill.mover.GraphMailClient", lambda settings: graph)

    result = move_message(make_settings(tmp_path), graph_id="old-id", destination="archive", execute=True)

    assert result["server_moved"] is True
    assert result["new_graph_id"] == "new-id"
    assert result["readback_verified"] is False
    assert result["local_state"] == "pending"
    assert result["completed"] is False
    assert store.updates == []


def test_move_allows_uncached_message_without_creating_local_row(monkeypatch, tmp_path):
    store = FakeStore(cached=False)
    graph = FakeGraph(None)
    monkeypatch.setattr("outlook_skill.mover.MailStore", lambda settings: store)
    monkeypatch.setattr("outlook_skill.mover.GraphMailClient", lambda settings: graph)

    result = move_message(make_settings(tmp_path), graph_id="old-id", destination="inbox", execute=True)

    assert result["readback_verified"] is True
    assert result["local_state"] == "not_cached"
    assert store.updates == []
