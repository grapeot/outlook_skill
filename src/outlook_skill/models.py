from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class DownloadedMessage:
    account: str
    folder: str
    uid: str
    uidvalidity: str | None
    message_id: str | None
    subject: str | None
    from_addr: str | None
    received_at: str | None
    internaldate: str | None
    size: int | None
    flags: list[str]
    mime_path: str
    sha256: str


@dataclass(frozen=True)
class SpamRule:
    id: int
    rule_name: str
    label: str
    rule_type: str
    pattern: str
    confidence: float
    enabled: bool
    created_at: str
    updated_at: str


@dataclass(frozen=True)
class TriageLabel:
    id: int
    message_uid: str
    folder: str
    label: str
    rule_id: int | None
    confidence: float | None
    labeled_at: str
