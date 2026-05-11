from __future__ import annotations

import re
from dataclasses import dataclass
from email import policy
from email.parser import BytesParser
from pathlib import Path

from .config import Settings
from .errors import OutlookSkillError
from .models import SpamRule
from .store import MailStore


SUPPORTED_RULE_TYPES = {"sender_exact", "sender_domain", "subject_regex", "body_regex"}


@dataclass(frozen=True)
class TriageMatch:
    uid: str
    folder: str
    label: str
    rule_id: int
    rule_name: str
    confidence: float
    subject: str | None
    from_addr: str | None


def add_rule(settings: Settings, *, rule_name: str, label: str, rule_type: str, pattern: str, confidence: float) -> dict[str, object]:
    normalized_type = rule_type.strip()
    if normalized_type not in SUPPORTED_RULE_TYPES:
        raise OutlookSkillError(f"Unsupported spam rule type: {rule_type}")
    store = MailStore(settings)
    try:
        rule = store.add_spam_rule(
            rule_name=rule_name,
            label=label,
            rule_type=normalized_type,
            pattern=pattern,
            confidence=confidence,
        )
        return rule_to_dict(rule)
    finally:
        store.close()


def list_rules(settings: Settings) -> dict[str, object]:
    store = MailStore(settings)
    try:
        rules = [rule_to_dict(rule) for rule in store.list_spam_rules(enabled_only=False)]
        return {"count": len(rules), "rules": rules}
    finally:
        store.close()


def apply_rules(settings: Settings, *, folder: str | None = None, limit: int = 100, label: str | None = None) -> dict[str, object]:
    store = MailStore(settings)
    try:
        rules = store.list_spam_rules(enabled_only=True)
        if label is not None:
            rules = [rule for rule in rules if rule_label(rule) == label]
        candidates = store.iter_messages_for_triage(folder=folder, limit=limit)
        matches: list[dict[str, object]] = []
        for candidate in candidates:
            for rule in rules:
                if message_matches_rule(candidate, rule):
                    match = store.save_triage_label(
                        message_uid=str(candidate["uid"]),
                        folder=str(candidate["folder"]),
                        label=rule_label(rule),
                        rule_id=rule.id,
                        confidence=rule.confidence,
                    )
                    matches.append(
                        {
                            "uid": match.message_uid,
                            "folder": match.folder,
                            "label": match.label,
                            "rule_id": match.rule_id,
                            "rule_name": rule.rule_name,
                            "confidence": match.confidence,
                            "subject": candidate.get("subject"),
                            "from_addr": candidate.get("from_addr"),
                        }
                    )
                    break
        return {
            "checked_count": len(candidates),
            "matched_count": len(matches),
            "label_filter": label,
            "matches": matches,
        }
    finally:
        store.close()


def list_labels(settings: Settings, *, label: str | None = None, limit: int = 100) -> dict[str, object]:
    store = MailStore(settings)
    try:
        labels = [
            {
                "id": item.id,
                "message_uid": item.message_uid,
                "folder": item.folder,
                "label": item.label,
                "rule_id": item.rule_id,
                "confidence": item.confidence,
                "labeled_at": item.labeled_at,
            }
            for item in store.list_triage_labels(label=label, limit=limit)
        ]
        return {"count": len(labels), "labels": labels}
    finally:
        store.close()


def message_matches_rule(candidate: dict[str, object], rule: SpamRule) -> bool:
    subject = str(candidate.get("subject") or "")
    from_addr = str(candidate.get("from_addr") or "")
    pattern = rule.pattern
    if rule.rule_type == "sender_exact":
        return from_addr.lower() == pattern.lower()
    if rule.rule_type == "sender_domain":
        return from_addr.lower().endswith("@" + pattern.lower())
    if rule.rule_type == "subject_regex":
        return re.search(pattern, subject, re.IGNORECASE) is not None
    if rule.rule_type == "body_regex":
        mime_path = candidate.get("mime_path")
        if not isinstance(mime_path, str):
            return False
        return re.search(pattern, load_body_text(Path(mime_path)), re.IGNORECASE) is not None
    return False


def load_body_text(mime_path: Path) -> str:
    message = BytesParser(policy=policy.default).parsebytes(mime_path.read_bytes())
    parts: list[str] = []
    for part in message.walk():
        if part.is_multipart():
            continue
        if part.get_content_disposition() == "attachment":
            continue
        if part.get_content_type() not in {"text/plain", "text/html"}:
            continue
        try:
            parts.append(str(part.get_content()))
        except Exception:
            payload = part.get_payload(decode=True) or b""
            payload_bytes = payload if isinstance(payload, bytes) else str(payload).encode("utf-8", errors="replace")
            charset = part.get_content_charset() or "utf-8"
            parts.append(payload_bytes.decode(charset, errors="replace"))
    return "\n".join(parts)


def rule_to_dict(rule: SpamRule) -> dict[str, object]:
    return {
        "id": rule.id,
        "rule_name": rule.rule_name,
        "label": rule_label(rule),
        "rule_type": rule.rule_type,
        "pattern": rule.pattern,
        "confidence": rule.confidence,
        "enabled": rule.enabled,
        "created_at": rule.created_at,
        "updated_at": rule.updated_at,
    }


def rule_label(rule: SpamRule) -> str:
    label = getattr(rule, "label", "spam")
    return label if isinstance(label, str) else "spam"
