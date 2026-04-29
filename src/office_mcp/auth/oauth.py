"""Thin wrapper over :mod:`msal` for the Azure AD authorization-code flow.

We deliberately keep the surface narrow: build an authorization URL,
exchange a code for tokens, and refresh a token. MSAL's broader
``acquire_token_silent`` / token-cache features are not used here — the
:class:`~office_mcp.auth.store.TokenStore` is the source of truth, so
we avoid the impedance mismatch between two caches.
"""

from __future__ import annotations

import time
from dataclasses import dataclass
from typing import Any

import msal
import structlog

from office_mcp.auth.store import TokenEntry

log = structlog.get_logger(__name__)

# Delegated scopes — match the API permissions granted on the app reg.
# `offline_access` is what unlocks the refresh token. `User.Read` lets
# us populate `oid` / `preferred_username` claims in the ID token.
SCOPES: tuple[str, ...] = ("Files.ReadWrite", "User.Read", "offline_access")


@dataclass(frozen=True, slots=True)
class OAuthConfig:
    """Static configuration for the Azure AD client."""

    tenant_id: str
    client_id: str
    client_secret: str
    redirect_uri: str

    @property
    def authority(self) -> str:
        return f"https://login.microsoftonline.com/{self.tenant_id}"


class OAuthError(Exception):
    """Raised when Azure AD rejects an exchange or refresh."""


class OAuthClient:
    """Driver for the authorization-code flow against Azure AD."""

    def __init__(self, config: OAuthConfig) -> None:
        self._config = config
        self._app = msal.ConfidentialClientApplication(
            client_id=config.client_id,
            client_credential=config.client_secret,
            authority=config.authority,
        )

    @property
    def redirect_uri(self) -> str:
        return self._config.redirect_uri

    def build_auth_url(self, state: str) -> str:
        """Return the URL the user should open to start the flow."""
        url = self._app.get_authorization_request_url(
            scopes=list(SCOPES),
            state=state,
            redirect_uri=self._config.redirect_uri,
        )
        # MSAL types the return as ``str`` already; this is just to
        # placate stricter consumers and document the contract.
        return str(url)

    def exchange_code(self, code: str) -> TokenEntry:
        """Trade an authorization ``code`` for a :class:`TokenEntry`."""
        result: dict[str, Any] = self._app.acquire_token_by_authorization_code(
            code=code,
            scopes=list(SCOPES),
            redirect_uri=self._config.redirect_uri,
        )
        return _to_entry_or_raise(result, action="exchange")

    def refresh(self, refresh_token: str) -> TokenEntry:
        """Trade a refresh token for a fresh :class:`TokenEntry`."""
        result: dict[str, Any] = self._app.acquire_token_by_refresh_token(
            refresh_token=refresh_token,
            scopes=list(SCOPES),
        )
        return _to_entry_or_raise(result, action="refresh")


def _to_entry_or_raise(result: dict[str, Any], *, action: str) -> TokenEntry:
    """Normalize MSAL's dict response into a :class:`TokenEntry` or raise."""
    if "access_token" not in result:
        # MSAL surfaces failures as ``{"error": "...", "error_description": "..."}``.
        description = result.get("error_description") or result.get("error") or "unknown error"
        log.warning("oauth.failure", action=action, error=str(description))
        raise OAuthError(f"oauth {action} failed: {description}")

    id_claims = result.get("id_token_claims") or {}
    return TokenEntry(
        access_token=str(result["access_token"]),
        refresh_token=result.get("refresh_token"),
        expires_at=time.time() + int(result.get("expires_in", 3600)),
        user_oid=str(id_claims.get("oid", "")),
        user_upn=str(id_claims.get("preferred_username", "")),
        user_email=id_claims.get("email"),
    )
