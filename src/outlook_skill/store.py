from __future__ import annotations

import hashlib
import sqlite3
from datetime import UTC, datetime
from pathlib import Path
from types import SimpleNamespace
from typing import cast

from .config import Settings
from .errors import OutlookSkillError
from .models import DownloadedMessage, SpamRule, TriageLabel


MESSAGES_SCHEMA = """
CREATE TABLE IF NOT EXISTS messages (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    account TEXT NOT NULL,
    folder TEXT NOT NULL,
    parent_folder_id TEXT,
    uid TEXT NOT NULL,
    uidvalidity TEXT,
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
    UNIQUE(account, uid)
);
"""

TRIAGE_LABELS_SCHEMA = """
CREATE TABLE IF NOT EXISTS triage_labels (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    message_uid TEXT NOT NULL,
    folder TEXT NOT NULL,
    label TEXT NOT NULL,
    rule_id INTEGER,
    confidence REAL,
    labeled_at TEXT NOT NULL,
    FOREIGN KEY(rule_id) REFERENCES spam_rules(id),
    UNIQUE(message_uid, label)
);
"""

SCHEMA = MESSAGES_SCHEMA + """
CREATE TABLE IF NOT EXISTS spam_rules (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    rule_name TEXT UNIQUE NOT NULL,
    label TEXT NOT NULL DEFAULT 'spam',
    rule_type TEXT NOT NULL,
    pattern TEXT NOT NULL,
    confidence REAL NOT NULL DEFAULT 0.5,
    enabled INTEGER NOT NULL DEFAULT 1,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);
""" + TRIAGE_LABELS_SCHEMA


class MailStore:
    def __init__(self, settings: Settings):
        self.settings = settings
        self.settings.messages_dir.mkdir(parents=True, exist_ok=True)
        self.settings.data_dir.mkdir(parents=True, exist_ok=True)
        self.connection = sqlite3.connect(self.settings.db_path)
        self.connection.row_factory = sqlite3.Row
        self.connection.executescript(SCHEMA)
        self._migrate_schema()
        self.connection.commit()

    def _migrate_schema(self) -> None:
        message_columns = {
            row[1]
            for row in self.connection.execute("PRAGMA table_info(messages)").fetchall()
        }
        if "parent_folder_id" not in message_columns:
            self.connection.close()
            raise OutlookSkillError(
                "This database uses the legacy folder-scoped message key. "
                "Run `mail migrate-current-state` before opening it with this version."
            )
        columns = {
            row[1]
            for row in self.connection.execute("PRAGMA table_info(spam_rules)").fetchall()
        }
        if "label" not in columns:
            self.connection.execute("ALTER TABLE spam_rules ADD COLUMN label TEXT NOT NULL DEFAULT 'spam'")

    def close(self) -> None:
        self.connection.close()

    def get_message(self, *, account: str, uid: str) -> dict[str, object] | None:
        row = self.connection.execute(
            "SELECT * FROM messages WHERE account = ? AND uid = ?",
            (account, uid),
        ).fetchone()
        return dict(row) if row is not None else None

    def has_message(self, *, account: str, uid: str) -> bool:
        return self.get_message(account=account, uid=uid) is not None

    def update_message_location(
        self,
        *,
        account: str,
        old_graph_id: str,
        new_graph_id: str,
        folder: str,
        parent_folder_id: str,
    ) -> dict[str, object]:
        current = self.get_message(account=account, uid=old_graph_id)
        if current is None:
            raise OutlookSkillError(f"Cannot reconcile missing local message: {old_graph_id}")
        if new_graph_id != old_graph_id and self.get_message(account=account, uid=new_graph_id) is not None:
            raise OutlookSkillError(f"Cannot reconcile Graph ID collision: {new_graph_id}")

        old_path = Path(str(current["mime_path"]))
        new_path = old_path
        renamed = False
        if new_graph_id != old_graph_id:
            new_path = self._mime_path(account=account, uid=new_graph_id)
            if new_path.exists() and new_path != old_path:
                raise OutlookSkillError(f"Cannot move MIME to an existing path: {new_path}")
            old_path.replace(new_path)
            renamed = True
        try:
            cursor = self.connection.execute(
                """
                UPDATE messages
                SET uid = ?, folder = ?, parent_folder_id = ?, mime_path = ?
                WHERE account = ? AND uid = ?
                """,
                (new_graph_id, folder, parent_folder_id, str(new_path), account, old_graph_id),
            )
            if cursor.rowcount != 1:
                raise OutlookSkillError(f"Local reconciliation did not update exactly one row: {old_graph_id}")
            self.connection.execute(
                "UPDATE triage_labels SET message_uid = ?, folder = ? WHERE message_uid = ?",
                (new_graph_id, folder, old_graph_id),
            )
            self.connection.commit()
        except Exception:
            self.connection.rollback()
            if renamed:
                new_path.replace(old_path)
            raise
        return {
            "old_graph_id": old_graph_id,
            "new_graph_id": new_graph_id,
            "folder": folder,
            "parent_folder_id": parent_folder_id,
            "mime_path": str(new_path),
        }

    def save_message(
        self,
        *,
        account: str,
        folder: str,
        uid: str,
        uidvalidity: str | None,
        raw_message: bytes,
        message_id: str | None,
        subject: str | None,
        from_addr: str | None,
        received_at: str | None,
        internaldate: str | None,
        size: int | None,
        flags_json: str,
        parent_folder_id: str | None = None,
    ) -> DownloadedMessage:
        sha256 = hashlib.sha256(raw_message).hexdigest()
        existing = self.get_message(account=account, uid=uid)
        mime_path = Path(str(existing["mime_path"])) if existing else self._mime_path(account=account, uid=uid)
        mime_path.parent.mkdir(parents=True, exist_ok=True)
        if not mime_path.exists() or mime_path.read_bytes() != raw_message:
            mime_path.write_bytes(raw_message)
        downloaded_at = datetime.now(UTC).isoformat()
        self.connection.execute(
            """
            INSERT INTO messages (
                account, folder, parent_folder_id, uid, uidvalidity, message_id, subject, from_addr,
                received_at, internaldate, size, flags_json, mime_path, downloaded_at, sha256
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(account, uid) DO UPDATE SET
                folder = excluded.folder,
                parent_folder_id = excluded.parent_folder_id,
                uidvalidity = excluded.uidvalidity,
                message_id = excluded.message_id,
                subject = excluded.subject,
                from_addr = excluded.from_addr,
                received_at = excluded.received_at,
                internaldate = excluded.internaldate,
                size = excluded.size,
                flags_json = excluded.flags_json,
                mime_path = excluded.mime_path,
                downloaded_at = excluded.downloaded_at,
                sha256 = excluded.sha256
            """,
            (
                account,
                folder,
                parent_folder_id,
                uid,
                uidvalidity,
                message_id,
                subject,
                from_addr,
                received_at,
                internaldate,
                size,
                flags_json,
                str(mime_path),
                downloaded_at,
                sha256,
            ),
        )
        self.connection.commit()
        return DownloadedMessage(
            account=account,
            folder=folder,
            uid=uid,
            uidvalidity=uidvalidity,
            message_id=message_id,
            subject=subject,
            from_addr=from_addr,
            received_at=received_at,
            internaldate=internaldate,
            size=size,
            flags=flags_json.split(",") if flags_json else [],
            mime_path=str(mime_path),
            sha256=sha256,
        )

    def list_local_messages(self, *, limit: int) -> list[dict[str, object]]:
        cursor = self.connection.execute(
            """
            SELECT account, folder, uid, uidvalidity, message_id, subject, from_addr, received_at,
                   internaldate, size, flags_json, mime_path, downloaded_at, sha256
            FROM messages
            ORDER BY COALESCE(received_at, downloaded_at) DESC
            LIMIT ?
            """,
            (limit,),
        )
        rows = cursor.fetchall()
        return [
            {
                "account": row[0],
                "folder": row[1],
                "uid": row[2],
                "uidvalidity": row[3],
                "message_id": row[4],
                "subject": row[5],
                "from_addr": row[6],
                "received_at": row[7],
                "internaldate": row[8],
                "size": row[9],
                "flags": row[10].split(",") if row[10] else [],
                "mime_path": row[11],
                "downloaded_at": row[12],
                "sha256": row[13],
            }
            for row in rows
        ]

    def find_messages(
        self,
        *,
        subject_substring: str | None = None,
        from_substring: str | None = None,
        graph_id: str | None = None,
    ) -> list[dict[str, object]]:
        query = (
            "SELECT account, folder, uid, uidvalidity, message_id, subject, from_addr, "
            "received_at, internaldate, size, flags_json, mime_path, downloaded_at, sha256 "
            "FROM messages WHERE 1=1"
        )
        params: list[object] = []
        if graph_id is not None:
            query += " AND uid = ?"
            params.append(graph_id)
        if subject_substring:
            query += " AND LOWER(COALESCE(subject, '')) LIKE ?"
            params.append(f"%{subject_substring.lower()}%")
        if from_substring:
            query += " AND LOWER(COALESCE(from_addr, '')) LIKE ?"
            params.append(f"%{from_substring.lower()}%")
        query += " ORDER BY COALESCE(received_at, downloaded_at) DESC"
        rows = self.connection.execute(query, tuple(params)).fetchall()
        return [
            {
                "account": row[0],
                "folder": row[1],
                "uid": row[2],
                "uidvalidity": row[3],
                "message_id": row[4],
                "subject": row[5],
                "from_addr": row[6],
                "received_at": row[7],
                "internaldate": row[8],
                "size": row[9],
                "flags": row[10].split(",") if row[10] else [],
                "mime_path": row[11],
                "downloaded_at": row[12],
                "sha256": row[13],
            }
            for row in rows
        ]

    def add_spam_rule(self, *, rule_name: str, label: str, rule_type: str, pattern: str, confidence: float) -> SpamRule:
        now = datetime.now(UTC).isoformat()
        cursor = self.connection.execute(
            """
            INSERT INTO spam_rules (rule_name, label, rule_type, pattern, confidence, enabled, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, 1, ?, ?)
            """,
            (rule_name, label, rule_type, pattern, confidence, now, now),
        )
        self.connection.commit()
        rule_id = cursor.lastrowid
        return _make_spam_rule(
            id=int(rule_id if rule_id is not None else 0),
            rule_name=rule_name,
            label=label,
            rule_type=rule_type,
            pattern=pattern,
            confidence=confidence,
            enabled=True,
            created_at=now,
            updated_at=now,
        )

    def list_spam_rules(self, *, enabled_only: bool = True) -> list[SpamRule]:
        query = "SELECT id, rule_name, label, rule_type, pattern, confidence, enabled, created_at, updated_at FROM spam_rules"
        params: tuple[object, ...] = ()
        if enabled_only:
            query += " WHERE enabled = 1"
        query += " ORDER BY id ASC"
        rows = self.connection.execute(query, params).fetchall()
        return [
            _make_spam_rule(
                id=row[0],
                rule_name=row[1],
                label=row[2],
                rule_type=row[3],
                pattern=row[4],
                confidence=row[5],
                enabled=bool(row[6]),
                created_at=row[7],
                updated_at=row[8],
            )
            for row in rows
        ]

    def save_triage_label(
        self,
        *,
        message_uid: str,
        folder: str,
        label: str,
        rule_id: int | None,
        confidence: float | None,
    ) -> TriageLabel:
        labeled_at = datetime.now(UTC).isoformat()
        self.connection.execute(
            """
            INSERT OR IGNORE INTO triage_labels (message_uid, folder, label, rule_id, confidence, labeled_at)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (message_uid, folder, label, rule_id, confidence, labeled_at),
        )
        self.connection.commit()
        row = self.connection.execute(
            """
                SELECT id, message_uid, folder, label, rule_id, confidence, labeled_at
                FROM triage_labels WHERE message_uid = ? AND label = ?
                """,
                (message_uid, label),

        ).fetchone()
        return TriageLabel(
            id=row[0],
            message_uid=row[1],
            folder=row[2],
            label=row[3],
            rule_id=row[4],
            confidence=row[5],
            labeled_at=row[6],
        )

    def list_triage_labels(self, *, label: str | None = None, limit: int = 100) -> list[TriageLabel]:
        query = "SELECT id, message_uid, folder, label, rule_id, confidence, labeled_at FROM triage_labels"
        params: tuple[object, ...]
        if label:
            query += " WHERE label = ?"
            params = (label,)
        else:
            params = ()
        query += " ORDER BY id DESC LIMIT ?"
        params = (*params, limit)
        rows = self.connection.execute(query, params).fetchall()
        return [
            TriageLabel(
                id=row[0],
                message_uid=row[1],
                folder=row[2],
                label=row[3],
                rule_id=row[4],
                confidence=row[5],
                labeled_at=row[6],
            )
            for row in rows
        ]

    def iter_messages_for_triage(self, *, folder: str | None = None, limit: int = 100) -> list[dict[str, object]]:
        query = """
            SELECT uid, folder, subject, from_addr, mime_path, received_at
            FROM messages
        """
        params: tuple[object, ...]
        if folder:
            query += " WHERE folder = ?"
            params = (folder,)
        else:
            params = ()
        query += " ORDER BY COALESCE(received_at, downloaded_at) DESC LIMIT ?"
        params = (*params, limit)
        rows = self.connection.execute(query, params).fetchall()
        return [
            {
                "uid": row[0],
                "folder": row[1],
                "subject": row[2],
                "from_addr": row[3],
                "mime_path": row[4],
                "received_at": row[5],
            }
            for row in rows
        ]

    def _mime_path(self, *, account: str, uid: str) -> Path:
        safe_account = account.replace("@", "_at_")
        uid_hash = hashlib.sha256(uid.encode("utf-8")).hexdigest()[:24]
        return self.settings.messages_dir / f"{safe_account}_{uid_hash}.eml"


def _make_spam_rule(**kwargs: object) -> SpamRule:
    return cast(SpamRule, cast(object, SimpleNamespace(**kwargs)))
