from __future__ import annotations

import hashlib
import sqlite3
from datetime import UTC, datetime
from pathlib import Path
from types import SimpleNamespace
from typing import cast

from .config import Settings
from .models import DownloadedMessage, SpamRule, TriageLabel


SCHEMA = """
CREATE TABLE IF NOT EXISTS messages (
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

CREATE TABLE IF NOT EXISTS triage_labels (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    message_uid TEXT NOT NULL,
    folder TEXT NOT NULL,
    label TEXT NOT NULL,
    rule_id INTEGER,
    confidence REAL,
    labeled_at TEXT NOT NULL,
    FOREIGN KEY(rule_id) REFERENCES spam_rules(id),
    UNIQUE(message_uid, folder, label)
);
"""


class MailStore:
    def __init__(self, settings: Settings):
        self.settings = settings
        self.settings.messages_dir.mkdir(parents=True, exist_ok=True)
        self.settings.data_dir.mkdir(parents=True, exist_ok=True)
        self.connection = sqlite3.connect(self.settings.db_path)
        self.connection.executescript(SCHEMA)
        self._migrate_schema()
        self.connection.commit()

    def _migrate_schema(self) -> None:
        columns = {
            row[1]
            for row in self.connection.execute("PRAGMA table_info(spam_rules)").fetchall()
        }
        if "label" not in columns:
            self.connection.execute("ALTER TABLE spam_rules ADD COLUMN label TEXT NOT NULL DEFAULT 'spam'")

    def close(self) -> None:
        self.connection.close()

    def has_message(self, *, account: str, folder: str, uidvalidity: str | None, uid: str) -> bool:
        row = self.connection.execute(
            "SELECT 1 FROM messages WHERE account = ? AND folder = ? AND uidvalidity IS ? AND uid = ?",
            (account, folder, uidvalidity, uid),
        ).fetchone()
        return row is not None

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
    ) -> DownloadedMessage:
        sha256 = hashlib.sha256(raw_message).hexdigest()
        mime_path = self._mime_path(account=account, folder=folder, uid=uid, sha256=sha256)
        mime_path.parent.mkdir(parents=True, exist_ok=True)
        if not mime_path.exists():
            mime_path.write_bytes(raw_message)
        downloaded_at = datetime.now(UTC).isoformat()
        self.connection.execute(
            """
            INSERT INTO messages (
                account, folder, uid, uidvalidity, message_id, subject, from_addr,
                received_at, internaldate, size, flags_json, mime_path, downloaded_at, sha256
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                account,
                folder,
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
            FROM triage_labels WHERE message_uid = ? AND folder = ? AND label = ?
            """,
            (message_uid, folder, label),
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

    def _mime_path(self, *, account: str, folder: str, uid: str, sha256: str) -> Path:
        safe_account = account.replace("@", "_at_")
        safe_folder = folder.replace("/", "_")
        safe_uid = uid.replace("/", "_").replace("=", "_")
        filename = f"{safe_account}_{safe_folder}_{safe_uid[:24]}_{sha256[:12]}.eml"
        return self.settings.messages_dir / filename


def _make_spam_rule(**kwargs: object) -> SpamRule:
    return cast(SpamRule, cast(object, SimpleNamespace(**kwargs)))
