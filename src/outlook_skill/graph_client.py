from __future__ import annotations

from collections.abc import Mapping
from datetime import UTC, datetime, timedelta

import httpx

from .auth import AuthManager
from .config import Settings
from .errors import GraphApiError


IMMUTABLE_ID_HEADER = 'IdType="ImmutableId"'
WELL_KNOWN_FOLDERS = {
    "INBOX": "inbox",
    "DRAFTS": "drafts",
    "SENT": "sentitems",
    "SENTITEMS": "sentitems",
    "DELETED": "deleteditems",
    "DELETEDITEMS": "deleteditems",
    "ARCHIVE": "archive",
}
QueryValue = str | int | float | bool | None


class GraphMailClient:
    def __init__(self, settings: Settings):
        self.settings = settings
        self.client = self._build_client()

    def _build_client(self) -> httpx.Client:
        token = AuthManager(self.settings).get_access_token()
        self.client = httpx.Client(
            base_url=self.settings.graph_base_url,
            timeout=httpx.Timeout(120.0, connect=10.0),
            headers={
                "Authorization": f"Bearer {token}",
                "Accept": "application/json",
                "Prefer": IMMUTABLE_ID_HEADER,
            },
        )
        return self.client

    def _refresh_client(self) -> None:
        self.client.close()
        self.client = self._build_client()

    def close(self) -> None:
        self.client.close()

    def resolve_folder(self, folder_name: str) -> tuple[str, str]:
        normalized = folder_name.strip().upper()
        well_known = WELL_KNOWN_FOLDERS.get(normalized)
        if well_known:
            payload = self._get(f"/me/mailFolders/{well_known}")
            return str(payload["id"]), str(payload.get("displayName") or folder_name)

        escaped = folder_name.replace("'", "''")
        payload = self._get(
            "/me/mailFolders",
            params={
                "$filter": f"displayName eq '{escaped}'",
                "$select": "id,displayName",
                "$top": 1,
            },
        )
        values = payload.get("value")
        if not isinstance(values, list) or not values:
            raise GraphApiError(f"Graph folder not found: {folder_name}")
        first = values[0]
        if not isinstance(first, Mapping):
            raise GraphApiError(f"Graph folder payload is invalid for: {folder_name}")
        return str(first["id"]), str(first.get("displayName") or folder_name)

    def list_recent_messages(self, *, folder_id: str, days: int, limit: int) -> list[dict[str, str | None]]:
        cutoff = (datetime.now(UTC) - timedelta(days=days)).isoformat().replace("+00:00", "Z")
        results: list[dict[str, str | None]] = []
        next_url: str | None = f"/me/mailFolders/{folder_id}/messages"
        params: dict[str, QueryValue] | None = {
            "$filter": f"receivedDateTime ge {cutoff}",
            "$select": "id,internetMessageId,receivedDateTime,parentFolderId",
            "$orderby": "receivedDateTime desc",
            "$top": min(limit, 100),
        }
        while next_url and len(results) < limit:
            payload = self._get(next_url, params=params)
            params = None
            values = payload.get("value")
            if isinstance(values, list):
                for value in values:
                    if not isinstance(value, Mapping):
                        continue
                    message_id = value.get("id")
                    if not isinstance(message_id, str):
                        continue
                    internet_message_id = value.get("internetMessageId")
                    received = value.get("receivedDateTime")
                    results.append(
                        {
                            "id": message_id,
                            "internet_message_id": internet_message_id if isinstance(internet_message_id, str) else None,
                            "received_at": received if isinstance(received, str) else None,
                        }
                    )
                    if len(results) >= limit:
                        break
            next_link = payload.get("@odata.nextLink")
            next_url = next_link if isinstance(next_link, str) else None
        return results

    def fetch_message_mime(self, *, message_id: str) -> bytes:
        response = self.client.get(f"/me/messages/{message_id}/$value", headers={"Prefer": IMMUTABLE_ID_HEADER})
        if response.status_code == 401:
            self._refresh_client()
            response = self.client.get(f"/me/messages/{message_id}/$value", headers={"Prefer": IMMUTABLE_ID_HEADER})
        if response.is_error:
            raise GraphApiError(
                f"Graph MIME download failed with status {response.status_code}",
                status_code=response.status_code,
                response_text=response.text,
            )
        return response.content

    def _get(self, path_or_url: str, *, params: Mapping[str, QueryValue] | None = None) -> dict[str, object]:
        response = self.client.get(path_or_url, params=params)
        if response.status_code == 401:
            self._refresh_client()
            response = self.client.get(path_or_url, params=params)
        if response.is_error:
            raise GraphApiError(
                f"Graph request failed with status {response.status_code}",
                status_code=response.status_code,
                response_text=response.text,
            )
        payload = response.json()
        if not isinstance(payload, dict):
            raise GraphApiError("Graph returned an unexpected payload shape.")
        return payload
