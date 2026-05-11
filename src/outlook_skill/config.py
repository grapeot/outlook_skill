from __future__ import annotations

import json
import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv

from .errors import ConfigError


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_FOLDERS = ("INBOX", "Archive")
DEFAULT_DATA_DIR = "data/mail"
DEFAULT_AUTHORITY = "https://login.microsoftonline.com/consumers"
DEFAULT_SCOPE = "Mail.Read offline_access"
DEFAULT_TOKEN_CACHE_PATH = "data/mail/oauth_token_cache.json"
DEFAULT_GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


@dataclass(frozen=True)
class Settings:
    email: str
    client_id: str
    authority: str
    scopes: tuple[str, ...]
    graph_base_url: str
    default_folders: tuple[str, ...]
    token_cache_path: Path
    data_dir: Path
    messages_dir: Path
    markdowns_dir: Path
    db_path: Path


def load_settings() -> Settings:
    load_dotenv(PROJECT_ROOT / ".env", override=False)
    email = os.getenv("OUTLOOK_EMAIL", "").strip()
    if not email:
        raise ConfigError("Missing OUTLOOK_EMAIL")
    client_id = os.getenv("OUTLOOK_CLIENT_ID", "").strip()
    if not client_id:
        raise ConfigError("Missing OUTLOOK_CLIENT_ID")

    data_dir = to_project_path(os.getenv("OUTLOOK_DATA_DIR", DEFAULT_DATA_DIR).strip() or DEFAULT_DATA_DIR)
    token_cache_path = to_project_path(os.getenv("OUTLOOK_TOKEN_CACHE_PATH", DEFAULT_TOKEN_CACHE_PATH).strip() or DEFAULT_TOKEN_CACHE_PATH)
    scopes_raw = os.getenv("OUTLOOK_GRAPH_SCOPE", DEFAULT_SCOPE).strip() or DEFAULT_SCOPE

    default_folders_raw = os.getenv("OUTLOOK_DEFAULT_FOLDERS", ",".join(DEFAULT_FOLDERS)).strip() or ",".join(DEFAULT_FOLDERS)
    default_folders = tuple(part.strip() for part in default_folders_raw.split(",") if part.strip())

    return Settings(
        email=email,
        client_id=client_id,
        authority=os.getenv("OUTLOOK_AUTHORITY", DEFAULT_AUTHORITY).strip() or DEFAULT_AUTHORITY,
        scopes=tuple(part for part in scopes_raw.split() if part),
        graph_base_url=os.getenv("OUTLOOK_GRAPH_BASE_URL", DEFAULT_GRAPH_BASE_URL).strip() or DEFAULT_GRAPH_BASE_URL,
        default_folders=default_folders,
        token_cache_path=token_cache_path,
        data_dir=data_dir,
        messages_dir=data_dir / "messages",
        markdowns_dir=data_dir / "markdowns",
        db_path=data_dir / "mail.db",
    )


def doctor_info() -> dict[str, object]:
    load_dotenv(PROJECT_ROOT / ".env", override=False)
    email = os.getenv("OUTLOOK_EMAIL", "").strip()
    client_id = os.getenv("OUTLOOK_CLIENT_ID", "").strip()
    data_dir = to_project_path(os.getenv("OUTLOOK_DATA_DIR", DEFAULT_DATA_DIR).strip() or DEFAULT_DATA_DIR)
    token_cache_path = to_project_path(os.getenv("OUTLOOK_TOKEN_CACHE_PATH", DEFAULT_TOKEN_CACHE_PATH).strip() or DEFAULT_TOKEN_CACHE_PATH)
    return {
        "email_configured": bool(email),
        "email_masked": mask_value(email),
        "client_id_configured": bool(client_id),
        "client_id_masked": mask_value(client_id),
        "authority": os.getenv("OUTLOOK_AUTHORITY", DEFAULT_AUTHORITY).strip() or DEFAULT_AUTHORITY,
        "scopes": [part for part in (os.getenv("OUTLOOK_GRAPH_SCOPE", DEFAULT_SCOPE).strip() or DEFAULT_SCOPE).split() if part],
        "graph_base_url": os.getenv("OUTLOOK_GRAPH_BASE_URL", DEFAULT_GRAPH_BASE_URL).strip() or DEFAULT_GRAPH_BASE_URL,
        "default_folders": [
            part.strip()
            for part in (os.getenv("OUTLOOK_DEFAULT_FOLDERS", ",".join(DEFAULT_FOLDERS)).strip() or ",".join(DEFAULT_FOLDERS)).split(",")
            if part.strip()
        ],
        "token_cache_path": str(token_cache_path),
        "token_cache_exists": token_cache_path.exists(),
        "data_dir": str(data_dir),
        "data_dir_exists": data_dir.exists(),
        "db_path": str(data_dir / "mail.db"),
        "db_exists": (data_dir / "mail.db").exists(),
        "project_root": str(PROJECT_ROOT),
    }


def to_project_path(raw_path: str) -> Path:
    path = Path(raw_path)
    if not path.is_absolute():
        path = PROJECT_ROOT / path
    return path


def mask_value(value: str) -> str | None:
    if not value:
        return None
    if "@" in value and len(value) > 9:
        return f"{value[:4]}...{value[-5:]}"
    if len(value) <= 8:
        return "*" * len(value)
    return f"{value[:4]}...{value[-4:]}"


def pretty_json(data: object) -> str:
    return json.dumps(data, ensure_ascii=False, indent=2, sort_keys=True)
