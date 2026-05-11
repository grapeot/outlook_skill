from __future__ import annotations

from .config import Settings
from .downloader import list_local_mail


def list_messages(settings: Settings, *, limit: int = 50) -> dict[str, object]:
    return list_local_mail(settings, limit=limit)
