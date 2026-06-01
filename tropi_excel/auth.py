"""Token providers for the Microsoft Graph Excel API.

Two implementations:
  BrokerTokenProvider   — production: calls an internal HTTP token broker.
  LocalTokenProvider    — local dev: silently reuses the MSAL token cache from
                          the excel-spike device-code flow.

A factory function ``get_token_provider()`` picks the right one based on the
environment.
"""
from __future__ import annotations

import os
import time
from typing import Protocol

import requests

from .errors import ExcelAuthError

GRAPH_SCOPES = ["Sites.ReadWrite.All", "User.Read"]


# ---------------------------------------------------------------------------
# Protocol — anything that can return a bearer token string.
# ---------------------------------------------------------------------------

class TokenProvider(Protocol):
    def get_token(self) -> str:
        """Return a valid bearer access token (no 'Bearer ' prefix)."""
        ...


# ---------------------------------------------------------------------------
# BrokerTokenProvider
# ---------------------------------------------------------------------------

class BrokerTokenProvider:
    """Fetch a delegated token from the internal token broker service.

    The broker exposes ``GET {base_url}/token`` which returns JSON::

        {"access_token": "...", "expires_at": 1234567890.0}

    The token is cached locally until ~60 s before ``expires_at``.

    Configuration (via env vars or constructor arguments):
      GRAPH_EXCEL_BROKER_URL      Base URL of the broker (no trailing slash).
      GRAPH_EXCEL_BROKER_API_KEY  API key sent as ``X-API-Key`` header.
    """

    def __init__(
        self,
        base_url: str | None = None,
        api_key: str | None = None,
    ) -> None:
        self._base_url = (base_url or os.getenv("GRAPH_EXCEL_BROKER_URL", "")).rstrip("/")
        self._api_key = api_key or os.getenv("GRAPH_EXCEL_BROKER_API_KEY", "")
        if not self._base_url:
            raise ExcelAuthError("GRAPH_EXCEL_BROKER_URL is not set.")
        if not self._api_key:
            raise ExcelAuthError("GRAPH_EXCEL_BROKER_API_KEY is not set.")

        self._token: str | None = None
        self._expires_at: float = 0.0

    def get_token(self) -> str:
        now = time.time()
        if self._token and now < self._expires_at - 60:
            return self._token

        resp = requests.get(
            f"{self._base_url}/token",
            headers={"X-API-Key": self._api_key},
            timeout=15,
        )
        if not resp.ok:
            raise ExcelAuthError(
                f"Token broker returned {resp.status_code}: {resp.text[:300]}"
            )
        data = resp.json()
        if "access_token" not in data:
            raise ExcelAuthError(
                f"Token broker response missing 'access_token': {data}"
            )
        self._token = data["access_token"]
        self._expires_at = float(data.get("expires_at", now + 3600))
        return self._token


# ---------------------------------------------------------------------------
# LocalTokenProvider  (local testing without the broker)
# ---------------------------------------------------------------------------

class LocalTokenProvider:
    """Silently acquire a token from a saved MSAL token cache.

    Designed for local development. Uses the ``token_cache.bin`` produced by
    ``excel-spike``'s device-code flow (``python spike.py init`` + ``run``).

    Configuration (via env vars or the .env in the spike directory):
      M365_TENANT_ID  Azure AD tenant ID.
      M365_CLIENT_ID  Azure AD application (client) ID.

    Args:
        cache_path: Path to the serialisable MSAL token cache file.
                    Defaults to ``~/dev/excel-spike/token_cache.bin``.
        env_path:   Path to a .env file to load tenant/client IDs from.
                    Defaults to ``~/dev/excel-spike/.env``.
    """

    _DEFAULT_CACHE = os.path.expanduser("~/dev/excel-spike/token_cache.bin")
    _DEFAULT_ENV = os.path.expanduser("~/dev/excel-spike/.env")

    def __init__(
        self,
        cache_path: str | None = None,
        env_path: str | None = None,
    ) -> None:
        import msal  # imported here so missing msal only fails at construction

        self._msal = msal
        self._cache_path = cache_path or self._DEFAULT_CACHE
        self._env = self._load_env(env_path or self._DEFAULT_ENV)
        self._tenant_id = self._env.get("M365_TENANT_ID") or os.getenv("M365_TENANT_ID", "")
        self._client_id = self._env.get("M365_CLIENT_ID") or os.getenv("M365_CLIENT_ID", "")

        if not self._tenant_id:
            raise ExcelAuthError("M365_TENANT_ID is not set (env or .env file).")
        if not self._client_id:
            raise ExcelAuthError("M365_CLIENT_ID is not set (env or .env file).")

        self._token: str | None = None
        self._expires_at: float = 0.0

    @staticmethod
    def _load_env(path: str) -> dict[str, str]:
        result: dict[str, str] = {}
        try:
            with open(path) as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith("#") or "=" not in line:
                        continue
                    k, v = line.split("=", 1)
                    result[k.strip()] = v.strip()
        except FileNotFoundError:
            pass
        return result

    def _load_cache(self) -> "msal.SerializableTokenCache":
        cache = self._msal.SerializableTokenCache()
        try:
            with open(self._cache_path) as f:
                cache.deserialize(f.read())
        except FileNotFoundError:
            pass
        return cache

    def _save_cache(self, cache: "msal.SerializableTokenCache") -> None:
        if cache.has_state_changed:
            with open(self._cache_path, "w") as f:
                f.write(cache.serialize())
            os.chmod(self._cache_path, 0o600)

    def get_token(self) -> str:
        now = time.time()
        if self._token and now < self._expires_at - 60:
            return self._token

        cache = self._load_cache()
        app = self._msal.PublicClientApplication(
            self._client_id,
            authority=f"https://login.microsoftonline.com/{self._tenant_id}",
            token_cache=cache,
        )
        accounts = app.get_accounts()
        if not accounts:
            raise ExcelAuthError(
                "No saved account in MSAL cache. "
                "Run: cd ~/dev/excel-spike && python spike.py init && python spike.py run"
            )
        result = app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])
        self._save_cache(cache)

        if not result or "access_token" not in result:
            raise ExcelAuthError(
                f"Silent token acquisition failed: "
                f"{(result or {}).get('error_description', 'unknown')}"
            )

        self._token = result["access_token"]
        expires_in = int(result.get("expires_in", 3600))
        self._expires_at = now + expires_in
        return self._token


# ---------------------------------------------------------------------------
# Factory
# ---------------------------------------------------------------------------

def get_token_provider() -> TokenProvider:
    """Return the appropriate TokenProvider based on environment variables.

    If ``GRAPH_EXCEL_BROKER_URL`` is set → ``BrokerTokenProvider``.
    Otherwise → ``LocalTokenProvider`` (requires a prior spike device-code login).
    """
    if os.getenv("GRAPH_EXCEL_BROKER_URL", "").strip():
        return BrokerTokenProvider()
    return LocalTokenProvider()
