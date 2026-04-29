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

# Delegated scopes — only the application-specific ones. MSAL Python
# *rejects* the OIDC reserved scopes (`openid`, `profile`, `offline_access`)
# when passed explicitly: confidential clients get them — and the
# associated refresh token — added implicitly. `User.Read` is what
# populates `oid` / `preferred_username` in the ID token.
SCOPES: tuple[str, ...] = ("Files.ReadWrite", "User.Read")

#: MSAL refuses these as explicit scope inputs (see msal.application
#: ``_decorate_scope``). Kept here so the assertion test can catch a
#: future regression at lint time rather than in production.
_MSAL_RESERVED_SCOPES: frozenset[str] = frozenset({"openid", "profile", "offline_access"})


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
    oid = id_claims.get("oid")
    upn = id_claims.get("preferred_username")
    if not oid or not upn:
        # Should never happen with `openid` + `User.Read` in scope, but
        # silently storing an entry with empty oid/upn would make
        # subsequent lookups match by accident — fail closed.
        log.warning(
            "oauth.missing_id_claims",
            action=action,
            claims_keys=sorted(id_claims),
        )
        raise OAuthError(f"oauth {action} succeeded but id_token lacked oid/upn")

    return TokenEntry(
        access_token=str(result["access_token"]),
        refresh_token=result.get("refresh_token"),
        expires_at=time.time() + int(result.get("expires_in", 3600)),
        user_oid=str(oid),
        user_upn=str(upn),
        user_email=id_claims.get("email"),
    )
