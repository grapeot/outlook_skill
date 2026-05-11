from pathlib import Path

from outlook_skill.config import Settings
from outlook_skill.spam_triage import message_matches_rule
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


def test_store_can_add_and_list_spam_rules(tmp_path: Path):
    settings = make_settings(tmp_path)
    store = MailStore(settings)
    try:
        rule = store.add_spam_rule(
            rule_name="newsletters",
            label="spam",
            rule_type="sender_domain",
            pattern="newsletter.example.com",
            confidence=0.9,
        )
        rules = store.list_spam_rules(enabled_only=False)

        assert rule.rule_name == "newsletters"
        assert rule.label == "spam"
        assert len(rules) == 1
        assert rules[0].pattern == "newsletter.example.com"
    finally:
        store.close()


def test_message_matches_sender_domain_rule(tmp_path: Path):
    settings = make_settings(tmp_path)
    store = MailStore(settings)
    try:
        rule = store.add_spam_rule(
            rule_name="sender-domain",
            label="spam",
            rule_type="sender_domain",
            pattern="promo.example.com",
            confidence=0.8,
        )
        candidate = {
            "uid": "1",
            "folder": "Inbox",
            "subject": "Sale",
            "from_addr": "offer@promo.example.com",
            "mime_path": str((tmp_path / "mail.eml")),
        }

        assert message_matches_rule(candidate, rule) is True
    finally:
        store.close()


def test_message_matches_body_regex_rule(tmp_path: Path):
    settings = make_settings(tmp_path)
    store = MailStore(settings)
    try:
        rule = store.add_spam_rule(
            rule_name="unsubscribe-body",
            label="spam",
            rule_type="body_regex",
            pattern="unsubscribe",
            confidence=0.7,
        )
        eml_path = tmp_path / "mail.eml"
        eml_path.write_text("Subject: Hello\n\nClick here to unsubscribe.", encoding="utf-8")
        candidate = {
            "uid": "2",
            "folder": "Inbox",
            "subject": "Hello",
            "from_addr": "news@example.com",
            "mime_path": str(eml_path),
        }

        assert message_matches_rule(candidate, rule) is True
    finally:
        store.close()
