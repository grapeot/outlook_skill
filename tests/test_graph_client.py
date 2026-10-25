import json
from types import SimpleNamespace

import httpx

from outlook_skill import graph_client
from outlook_skill.graph_client import WELL_KNOWN_FOLDERS, GraphMailClient


def test_well_known_folders_maps_inbox():
    assert WELL_KNOWN_FOLDERS["INBOX"] == "inbox"


def test_move_posts_destination_and_uses_immutable_headers(monkeypatch):
    requests = []

    class Transport(httpx.BaseTransport):
        def handle_request(self, request):
            requests.append(request)
            if request.method == "POST" and request.url.path == "/v1.0/me/messages/old-id/move":
                assert json.loads(request.content) == {"destinationId": "archive-id"}
                return httpx.Response(201, json={"id": "new-id"})
            return httpx.Response(404)

    class FakeAuth:
        def __init__(self, settings):
            pass

        def get_access_token(self):
            return "TOKEN"

    original = httpx.Client
    monkeypatch.setattr(graph_client, "AuthManager", FakeAuth)
    monkeypatch.setattr(
        graph_client.httpx,
        "Client",
        lambda *args, **kwargs: original(*args, transport=Transport(), **kwargs),
    )
    client = GraphMailClient(SimpleNamespace(graph_base_url="https://example.test/v1.0"))
    try:
        result = client.move_message(message_id="old-id", destination_folder_id="archive-id")
    finally:
        client.close()

    assert result["id"] == "new-id"
    assert requests[0].headers["Prefer"] == 'IdType="ImmutableId"'
