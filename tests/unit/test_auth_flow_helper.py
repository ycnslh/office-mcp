"""Unit tests for the OAuth flow helper used by tools.

The helper sits between :class:`~office_mcp.utils.errors.AuthRequiredError`
and the LLM-facing response: it mints a ``state``, registers a pending
claim that carries the MCP session id, and returns the Microsoft
authorize URL.
"""

from __future__ import annotations

from unittest.mock import MagicMock

import pytest

from office_mcp.auth.flow import auth_required_response, start_oauth_flow
from office_mcp.auth.oauth import OAuthClient, OAuthConfig
from office_mcp.auth.store import TokenStore
from office_mcp.config import Settings
from office_mcp.runtime import set_oauth_client, set_settings, set_token_store

CONFIG = OAuthConfig(
    tenant_id="tenant-id",
    client_id="client-id",
    client_secret="secret",  # noqa: S106 — dummy literal for tests
    redirect_uri="http://localhost:8080/auth/callback",
)


@pytest.fixture
def wired_runtime(monkeypatch: pytest.MonkeyPatch) -> tuple[TokenStore, MagicMock]:
    """Build an OAuthClient with msal mocked, plus a fresh TokenStore."""
    fake_msal_app = MagicMock()
    fake_msal_app.get_authorization_request_url.side_effect = (
        lambda **kwargs: f"https://login.microsoftonline.com/authorize?state={kwargs['state']}"
    )
    monkeypatch.setattr(
        "office_mcp.auth.oauth.msal.ConfidentialClientApplication",
        MagicMock(return_value=fake_msal_app),
    )

    store = TokenStore()
    set_oauth_client(OAuthClient(CONFIG))
    set_token_store(store)
    set_settings(Settings(public_base_url="https://office-mcp.test"))
    return store, fake_msal_app


def test_start_oauth_flow_generates_state_and_url(
    wired_runtime: tuple[TokenStore, MagicMock],
) -> None:
    store, _ = wired_runtime

    state, url = start_oauth_flow(mcp_session_id="session-1")

    assert len(state) >= 32  # secrets.token_urlsafe(32) → ≥ 43 chars
    assert "login.microsoftonline.com" in url
    assert f"state={state}" in url


def test_start_oauth_flow_records_pending_with_session(
    wired_runtime: tuple[TokenStore, MagicMock],
) -> None:
    store, _ = wired_runtime

    state, _ = start_oauth_flow(mcp_session_id="session-1")
    pending = store.consume_pending(state)

    assert pending is not None
    assert pending.mcp_session_id == "session-1"


def test_start_oauth_flow_accepts_none_session(
    wired_runtime: tuple[TokenStore, MagicMock],
) -> None:
    """The /auth/login endpoint also calls into start-flow-like code
    without a session id; the helper must tolerate that."""
    store, _ = wired_runtime

    state, _ = start_oauth_flow(mcp_session_id=None)
    pending = store.consume_pending(state)

    assert pending is not None
    assert pending.mcp_session_id is None


def test_start_oauth_flow_raises_when_oauth_disabled() -> None:
    set_oauth_client(None)
    set_token_store(TokenStore())

    with pytest.raises(RuntimeError, match="OAuth"):
        start_oauth_flow(mcp_session_id="session-1")


def test_auth_required_response_shape(
    wired_runtime: tuple[TokenStore, MagicMock],
) -> None:
    """``login_url`` is the *short, our-domain* redirect — not the full
    Microsoft URL — so the LLM has minimal surface to mistype when
    surfacing it to the user. The redirect endpoint
    (:mod:`office_mcp.auth.routes._auth_start`) rebuilds the upstream
    URL server-side."""
    response = auth_required_response(mcp_session_id="session-1")

    assert response["auth_required"] is True
    assert response["login_url"].startswith("https://office-mcp.test/auth/start/")
    # Markdown-link rendering on its own line so the LLM can't truncate
    # the URL by paraphrasing it inline.
    assert f"[Connectez-vous ici]({response['login_url']})" in response["summary"]


def test_auth_required_response_state_in_url_matches_pending_claim(
    wired_runtime: tuple[TokenStore, MagicMock],
) -> None:
    """The state baked into the short URL must be the same one that
    /auth/start/{state} will look up — otherwise the redirect 404s."""
    store, _ = wired_runtime
    response = auth_required_response(mcp_session_id="session-1")

    state = response["login_url"].rsplit("/", 1)[-1]
    pending = store.peek_pending(state)

    assert pending is not None
    assert pending.mcp_session_id == "session-1"
