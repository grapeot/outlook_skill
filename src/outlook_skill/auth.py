from __future__ import annotations

from collections.abc import Mapping

import msal

from .config import Settings
from .errors import AuthRequiredError, OutlookSkillError


RESERVED_SCOPES = {"offline_access", "openid", "profile"}


class AuthManager:
    def __init__(self, settings: Settings):
        self.settings = settings
        self.cache = msal.SerializableTokenCache()
        if settings.token_cache_path.exists():
            self.cache.deserialize(settings.token_cache_path.read_text(encoding="utf-8"))
        self.app = msal.PublicClientApplication(
            client_id=settings.client_id,
            authority=settings.authority,
            token_cache=self.cache,
        )

    def login(self, *, device_code: bool = False) -> dict[str, object]:
        runtime_scopes = list(filter_reserved_scopes(self.settings.scopes))
        if device_code:
            flow = self.app.initiate_device_flow(scopes=runtime_scopes)
            message = flow.get("message")
            if not isinstance(message, str):
                raise OutlookSkillError("Failed to start device code flow.")
            result = self.app.acquire_token_by_device_flow(flow)
            if "access_token" not in result:
                raise OutlookSkillError(str(result.get("error_description") or result.get("error") or "Authentication failed."))
            self._persist_cache()
            scope_value = result.get("scope")
            return {
                "account": extract_account(result),
                "scopes": scope_value.split() if isinstance(scope_value, str) else [],
                "device_code_message": message,
                "token_cache_path": str(self.settings.token_cache_path),
            }

        result = self.app.acquire_token_interactive(
            scopes=runtime_scopes,
            prompt="select_account",
            port=0,
        )
        if "access_token" not in result:
            raise OutlookSkillError(str(result.get("error_description") or result.get("error") or "Authentication failed."))
        self._persist_cache()
        scope_value = result.get("scope")
        return {
            "account": extract_account(result),
            "scopes": scope_value.split() if isinstance(scope_value, str) else [],
            "token_cache_path": str(self.settings.token_cache_path),
        }

    def get_access_token(self) -> str:
        runtime_scopes = list(filter_reserved_scopes(self.settings.scopes))
        accounts = self.app.get_accounts(username=self.settings.email)
        if not accounts:
            accounts = self.app.get_accounts()
        if not accounts:
            raise AuthRequiredError("No cached OAuth account found. Run auth login first.")
        result = self.app.acquire_token_silent(runtime_scopes, account=accounts[0])
        if not result or "access_token" not in result:
            raise AuthRequiredError("OAuth token is missing or expired. Run auth login again.")
        self._persist_cache()
        token = result.get("access_token")
        if not isinstance(token, str):
            raise AuthRequiredError("OAuth token is invalid. Run auth login again.")
        return token

    def status(self) -> dict[str, object]:
        accounts = self.app.get_accounts(username=self.settings.email)
        if not accounts:
            accounts = self.app.get_accounts()
        return {
            "client_id_configured": True,
            "token_cache_path": str(self.settings.token_cache_path),
            "token_cache_exists": self.settings.token_cache_path.exists(),
            "cached_accounts": [account.get("username") for account in accounts if isinstance(account.get("username"), str)],
            "auth_ready": bool(accounts),
        }

    def _persist_cache(self) -> None:
        if self.cache.has_state_changed:
            self.settings.token_cache_path.parent.mkdir(parents=True, exist_ok=True)
            self.settings.token_cache_path.write_text(self.cache.serialize(), encoding="utf-8")


def extract_account(result: Mapping[str, object]) -> str | None:
    claims = result.get("id_token_claims")
    if isinstance(claims, Mapping):
        preferred = claims.get("preferred_username")
        if isinstance(preferred, str):
            return preferred
    account = result.get("account")
    if isinstance(account, Mapping):
        username = account.get("username")
        if isinstance(username, str):
            return username
    return None


def filter_reserved_scopes(scopes: tuple[str, ...]) -> tuple[str, ...]:
    return tuple(scope for scope in scopes if scope not in RESERVED_SCOPES)
