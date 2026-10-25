from __future__ import annotations

import httpx

from .config import Settings
from .errors import GraphApiError
from .graph_client import GraphMailClient
from .store import MailStore


def move_message(
    settings: Settings,
    *,
    graph_id: str,
    destination: str,
    execute: bool = False,
) -> dict[str, object]:
    if destination not in ("archive", "inbox"):
        raise ValueError(f"Unsupported move destination: {destination}")
    operation = "archive" if destination == "archive" else "unarchive"
    if not execute:
        return {
            "operation": operation,
            "graph_id": graph_id,
            "destination": destination,
            "execute": False,
            "server_moved": False,
            "readback_verified": False,
            "local_state": "not_run",
            "completed": False,
        }

    store = MailStore(settings)
    client = GraphMailClient(settings)
    try:
        cached = store.get_message(account=settings.email, uid=graph_id) is not None
        destination_folder_id, destination_name = client.resolve_folder(destination)
        client.get_message_metadata(message_id=graph_id)
        moved = client.move_message(message_id=graph_id, destination_folder_id=destination_folder_id)
        current_id = str(moved["id"])
        result: dict[str, object] = {
            "operation": operation,
            "graph_id": graph_id,
            "old_graph_id": graph_id,
            "new_graph_id": current_id,
            "graph_id_changed": current_id != graph_id,
            "destination": destination,
            "execute": True,
            "server_moved": True,
        }
        try:
            current = client.get_message_metadata(message_id=current_id)
        except (GraphApiError, httpx.RequestError):
            result.update(
                readback_verified=False,
                local_state="pending" if cached else "not_cached",
                completed=False,
            )
            return result
        if current.get("parentFolderId") != destination_folder_id:
            result.update(
                readback_verified=False,
                local_state="pending" if cached else "not_cached",
                completed=False,
                actual_parent_folder_id=current.get("parentFolderId"),
            )
            return result

        if cached:
            store.update_message_location(
                account=settings.email,
                old_graph_id=graph_id,
                new_graph_id=current_id,
                folder=destination_name,
                parent_folder_id=destination_folder_id,
            )
            local_state = "reconciled"
        else:
            local_state = "not_cached"
        result.update(readback_verified=True, local_state=local_state, completed=True)
        return result
    finally:
        client.close()
        store.close()
