from __future__ import annotations

import hashlib
import sqlite3
from collections import defaultdict
from pathlib import Path

from .config import Settings
from .downloader import extract_metadata
from .graph_client import GraphMailClient
from .store import MESSAGES_SCHEMA, TRIAGE_LABELS_SCHEMA


def migrate_current_state(settings: Settings, *, execute: bool = False) -> dict[str, object]:
    connection = (
        sqlite3.connect(settings.db_path)
        if execute
        else sqlite3.connect(f"{settings.db_path.resolve().as_uri()}?mode=ro", uri=True)
    )
    connection.row_factory = sqlite3.Row
    graph: GraphMailClient | None = None
    created_paths: list[Path] = []
    try:
        columns = {row[1] for row in connection.execute("PRAGMA table_info(messages)")}
        if "parent_folder_id" in columns:
            return {"already_migrated": True, "execute": False, "server_moved": False}

        rows = [dict(row) for row in connection.execute("SELECT * FROM messages ORDER BY id")]
        groups: dict[tuple[str, str], list[dict[str, object]]] = defaultdict(list)
        for row in rows:
            groups[(str(row["account"]), str(row["uid"]))].append(row)
        duplicates = [group for group in groups.values() if len(group) > 1]
        different_content = sum(1 for group in duplicates if len({row["sha256"] for row in group}) > 1)
        result: dict[str, object] = {
            "already_migrated": False,
            "execute": execute,
            "source_rows": len(rows),
            "canonical_rows": len(groups),
            "duplicate_groups": len(duplicates),
            "duplicate_groups_with_different_content": different_content,
            "triage_labels": connection.execute("SELECT count(*) FROM triage_labels").fetchone()[0],
        }
        if not execute:
            result["unreferenced_mime_count"] = len(rows) - len(groups)
            return result

        canonical: dict[tuple[str, str], dict[str, object]] = {}
        resolved_ids: dict[tuple[str, str], str] = {}
        unreferenced_paths: set[str] = set()
        if duplicates:
            graph = GraphMailClient(settings)
        for group in duplicates:
            old_id = str(group[0]["uid"])
            current = graph.get_message_metadata(message_id=old_id) if graph else None
            if current is None or not isinstance(current.get("id"), str):
                raise RuntimeError(f"Could not resolve current Graph message: {old_id}")
            current_id = str(current["id"])
            parent_folder_id = current.get("parentFolderId")
            if not isinstance(parent_folder_id, str):
                raise RuntimeError(f"Current Graph message has no parentFolderId: {old_id}")
            folder_payload = graph.get_folder(folder_id=parent_folder_id) if graph else {}
            folder = str(folder_payload.get("displayName") or "")
            if not folder:
                raise RuntimeError(f"Could not resolve current folder for: {old_id}")

            candidates = [
                row for row in group
                if str(row["uid"]) == current_id and str(row["folder"]).casefold() == folder.casefold()
            ] or [row for row in group if str(row["folder"]).casefold() == folder.casefold()]
            selected = max(candidates or group, key=lambda row: int(row["id"])).copy()
            selected["uid"] = current_id
            selected["folder"] = folder
            selected["parent_folder_id"] = parent_folder_id

            if len({row["sha256"] for row in group}) > 1:
                raw = graph.fetch_message_mime(message_id=current_id) if graph else b""
                metadata = extract_metadata(raw)
                sha256 = hashlib.sha256(raw).hexdigest()
                identity_hash = hashlib.sha256(
                    f"{selected['account']}\0{current_id}".encode("utf-8")
                ).hexdigest()[:24]
                path = settings.messages_dir / f"migration_{identity_hash}_{sha256[:12]}.eml"
                path.parent.mkdir(parents=True, exist_ok=True)
                if path.exists():
                    if path.read_bytes() != raw:
                        raise RuntimeError(f"Migration MIME path already contains different content: {path}")
                else:
                    path.write_bytes(raw)
                    created_paths.append(path)
                selected.update(
                    message_id=metadata["message_id"],
                    subject=metadata["subject"],
                    from_addr=metadata["from_addr"],
                    received_at=metadata["received_at"] or selected["received_at"],
                    size=len(raw),
                    sha256=sha256,
                    mime_path=str(path),
                )

            key = (str(selected["account"]), current_id)
            if key in canonical:
                raise RuntimeError(f"Current Graph ID maps to multiple legacy groups: {current_id}")
            canonical[key] = selected
            resolved_ids[(str(group[0]["account"]), old_id)] = current_id
            unreferenced_paths.update(str(row["mime_path"]) for row in group)

        for key, group in groups.items():
            if len(group) == 1:
                if key in canonical:
                    raise RuntimeError(f"Current Graph ID collides with an existing singleton: {key[1]}")
                row = group[0].copy()
                row["parent_folder_id"] = None
                canonical[key] = row
        unreferenced_paths.difference_update(str(row["mime_path"]) for row in canonical.values())
        result["unreferenced_mime_count"] = len(unreferenced_paths)

        id_aliases = {
            str(row["uid"]): resolved_ids[(str(group[0]["account"]), str(group[0]["uid"]))]
            for group in duplicates for row in group
        }
        message_values = [
            (
                row["id"], row["account"], row["folder"], row["parent_folder_id"], row["uid"],
                row["uidvalidity"], row["message_id"], row["subject"], row["from_addr"],
                row["received_at"], row["internaldate"], row["size"], row["flags_json"],
                row["mime_path"], row["downloaded_at"], row["sha256"],
            )
            for row in canonical.values()
        ]
        try:
            connection.execute("BEGIN IMMEDIATE")
            connection.execute(MESSAGES_SCHEMA.replace("CREATE TABLE IF NOT EXISTS messages", "CREATE TABLE messages_v2"))
            connection.executemany(
                """
                INSERT INTO messages_v2 (
                    id, account, folder, parent_folder_id, uid, uidvalidity, message_id,
                    subject, from_addr, received_at, internaldate, size, flags_json,
                    mime_path, downloaded_at, sha256
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                message_values,
            )
            connection.execute(
                TRIAGE_LABELS_SCHEMA.replace(
                    "CREATE TABLE IF NOT EXISTS triage_labels", "CREATE TABLE triage_labels_v2"
                )
            )
            label_rows = connection.execute(
                "SELECT * FROM triage_labels ORDER BY id"
            ).fetchall()
            latest_labels: dict[tuple[str, str], sqlite3.Row] = {}
            for row in label_rows:
                latest_labels[(str(row["message_uid"]), str(row["label"]))] = row
            connection.executemany(
                """
                INSERT INTO triage_labels_v2 (
                    id, message_uid, folder, label, rule_id, confidence, labeled_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                [
                    (
                        row["id"], id_aliases.get(str(row["message_uid"]), str(row["message_uid"])),
                        row["folder"], row["label"], row["rule_id"], row["confidence"], row["labeled_at"],
                    )
                    for row in latest_labels.values()
                ],
            )
            connection.execute("DROP TABLE messages")
            connection.execute("ALTER TABLE messages_v2 RENAME TO messages")
            connection.execute("DROP TABLE triage_labels")
            connection.execute("ALTER TABLE triage_labels_v2 RENAME TO triage_labels")
            connection.commit()
        except Exception:
            connection.rollback()
            raise
        result["migrated"] = True
        return result
    except Exception:
        for path in created_paths:
            path.unlink(missing_ok=True)
        raise
    finally:
        if graph is not None:
            graph.close()
        connection.close()
