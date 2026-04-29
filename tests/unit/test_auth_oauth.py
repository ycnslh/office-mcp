"""Unit tests for the MSAL wrapper.

We patch :class:`msal.ConfidentialClientApplication` at the module level
so the OAuthClient never reaches Azure AD. The mocked instance lets us
inspect the arguments :mod:`msal` would have been called with and to
script its return values.
"""

from __future__ import annotations

import time
from unittest.mock import MagicMock

import pytest

from office_mcp.auth.oauth import SCOPES, OAuthClient, OAuthConfig, OAuthError

CONFIG = OAuthConfig(
    tenant_id="tenant-id",
    client_id="client-id",
    client_secret="secret",  # noqa: S106 — dummy literal for tests
    redirect_uri="http://localhost:8080/auth/callback",
)


@pytest.fixture
def fake_msal_app(monkeypatch: pytest.MonkeyPatch) -> MagicMock:
    fake_instance = MagicMock()
    fake_class = MagicMock(return_value=fake_instance)
    monkeypatch.setattr(
        "office_mcp.auth.oauth.msal.ConfidentialClientApplication",
        fake_class,
    )
    return fake_instance


def test_authority_uses_tenant_id() -> None:
    assert CONFIG.authority == "https://login.microsoftonline.com/tenant-id"


def test_build_auth_url_passes_state_and_redirect(fake_msal_app: MagicMock) -> None:
    fake_msal_app.get_authorization_request_url.return_value = "https://login/authorize"
    client = OAuthClient(CONFIG)

    url = client.build_auth_url(state="claim-K7M2X9")

    assert url == "https://login/authorize"
    fake_msal_app.get_authorization_request_url.assert_called_once()
    kwargs = fake_msal_app.get_authorization_request_url.call_args.kwargs
    assert kwargs["state"] == "claim-K7M2X9"
    assert kwargs["redirect_uri"] == CONFIG.redirect_uri
    # All declared scopes are passed through.
    for scope in SCOPES:
        assert scope in kwargs["scopes"]


def test_exchange_code_returns_token_entry(fake_msal_app: MagicMock) -> None:
    fake_msal_app.acquire_token_by_authorization_code.return_value = {
        "access_token": "AT",
        "refresh_token": "RT",
        "expires_in": 3600,
        "id_token_claims": {
            "oid": "oid-1",
            "preferred_username": "alice@contoso.com",
            "email": "alice@contoso.com",
        },
    }
    client = OAuthClient(CONFIG)

    entry = client.exchange_code(code="auth-code")

    assert entry.access_token == "AT"
    assert entry.refresh_token == "RT"
    assert entry.user_oid == "oid-1"
    assert entry.user_upn == "alice@contoso.com"
    assert entry.user_email == "alice@contoso.com"
    assert entry.expires_at > time.time() + 3500


def test_exchange_code_raises_on_msal_error(fake_msal_app: MagicMock) -> None:
    fake_msal_app.acquire_token_by_authorization_code.return_value = {
        "error": "invalid_grant",
        "error_description": "code expired",
    }
    client = OAuthClient(CONFIG)

    with pytest.raises(OAuthError, match="code expired"):
        client.exchange_code(code="bad")


def test_exchange_code_falls_back_to_error_when_description_missing(
    fake_msal_app: MagicMock,
) -> None:
    fake_msal_app.acquire_token_by_authorization_code.return_value = {
        "error": "invalid_request",
    }
    client = OAuthClient(CONFIG)

    with pytest.raises(OAuthError, match="invalid_request"):
        client.exchange_code(code="bad")


def test_refresh_returns_token_entry(fake_msal_app: MagicMock) -> None:
    fake_msal_app.acquire_token_by_refresh_token.return_value = {
        "access_token": "AT2",
        "refresh_token": "RT2",
        "expires_in": 3600,
        "id_token_claims": {
            "oid": "oid-1",
            "preferred_username": "alice@contoso.com",
        },
    }
    client = OAuthClient(CONFIG)

    entry = client.refresh(refresh_token="old-rt")

    assert entry.access_token == "AT2"
    assert entry.refresh_token == "RT2"
    fake_msal_app.acquire_token_by_refresh_token.assert_called_once()


def test_refresh_raises_on_msal_error(fake_msal_app: MagicMock) -> None:
    fake_msal_app.acquire_token_by_refresh_token.return_value = {
        "error": "interaction_required",
        "error_description": "consent revoked",
    }
    client = OAuthClient(CONFIG)

    with pytest.raises(OAuthError, match="consent revoked"):
        client.refresh(refresh_token="rt")


def test_redirect_uri_property_exposes_config_value(fake_msal_app: MagicMock) -> None:
    client = OAuthClient(CONFIG)
    assert client.redirect_uri == CONFIG.redirect_uri


def test_exchange_code_raises_when_id_token_lacks_oid(fake_msal_app: MagicMock) -> None:
    """An access token without identity claims is unusable downstream."""
    fake_msal_app.acquire_token_by_authorization_code.return_value = {
        "access_token": "AT",
        "refresh_token": "RT",
        "expires_in": 3600,
        "id_token_claims": {"preferred_username": "alice@contoso.com"},  # missing oid
    }
    client = OAuthClient(CONFIG)

    with pytest.raises(OAuthError, match="oid"):
        client.exchange_code(code="code")


def test_exchange_code_raises_when_id_token_lacks_upn(fake_msal_app: MagicMock) -> None:
    fake_msal_app.acquire_token_by_authorization_code.return_value = {
        "access_token": "AT",
        "refresh_token": "RT",
        "expires_in": 3600,
        "id_token_claims": {"oid": "oid-1"},  # missing preferred_username
    }
    client = OAuthClient(CONFIG)

    with pytest.raises(OAuthError, match="upn"):
        client.exchange_code(code="code")
