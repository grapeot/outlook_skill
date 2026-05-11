from pathlib import Path

import pytest

from outlook_skill.config import doctor_info, load_settings
from outlook_skill.errors import ConfigError


def test_load_settings_reads_graph_environment(monkeypatch, tmp_path):
    monkeypatch.setenv("OUTLOOK_EMAIL", "duck@example.com")
    monkeypatch.setenv("OUTLOOK_CLIENT_ID", "client-id")
    monkeypatch.setenv("OUTLOOK_DATA_DIR", str(tmp_path / "mail"))
    monkeypatch.setenv("OUTLOOK_TOKEN_CACHE_PATH", str(tmp_path / "mail" / "oauth_token_cache.json"))
    monkeypatch.setenv("OUTLOOK_GRAPH_SCOPE", "Mail.Read offline_access")

    settings = load_settings()

    assert settings.email == "duck@example.com"
    assert settings.client_id == "client-id"
    assert settings.token_cache_path == tmp_path / "mail" / "oauth_token_cache.json"
    assert settings.scopes == ("Mail.Read", "offline_access")


def test_load_settings_requires_client_id(monkeypatch):
    monkeypatch.setenv("OUTLOOK_EMAIL", "duck@example.com")
    monkeypatch.setenv("OUTLOOK_CLIENT_ID", "")

    with pytest.raises(ConfigError):
        load_settings()


def test_doctor_info_reports_graph_fields(monkeypatch, tmp_path):
    cache_path = tmp_path / "mail" / "oauth_token_cache.json"
    cache_path.parent.mkdir(parents=True)
    cache_path.write_text("{}", encoding="utf-8")
    monkeypatch.setenv("OUTLOOK_EMAIL", "duck@example.com")
    monkeypatch.setenv("OUTLOOK_CLIENT_ID", "abcd1234wxyz")
    monkeypatch.setenv("OUTLOOK_TOKEN_CACHE_PATH", str(cache_path))
    monkeypatch.setenv("OUTLOOK_DATA_DIR", str(tmp_path / "mail"))

    info = doctor_info()

    assert info["email_masked"] == "duck...e.com"
    assert info["client_id_configured"] is True
    assert info["token_cache_exists"] is True
