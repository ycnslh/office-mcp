"""Unit tests for the auth_required path in the tool wrapper.

The :func:`office_mcp.instrumentation.instrumented` wrapper must catch
:class:`~office_mcp.utils.errors.AuthRequiredError` and convert it to a
structured response — otherwise every Graph-touching tool would
surface a raw exception to the LLM instead of the URL the user needs
to click.
"""

from __future__ import annotations

from unittest.mock import MagicMock

import pytest

from office_mcp.auth.oauth import OAuthClient, OAuthConfig
from office_mcp.auth.store import TokenStore
from office_mcp.context import bind_mcp_session_id
from office_mcp.instrumentation import instrumented
from office_mcp.runtime import set_oauth_client, set_token_store
from office_mcp.utils.errors import AuthRequiredError

CONFIG = OAuthConfig(
    tenant_id="tenant-id",
    client_id="client-id",
    client_secret="secret",  # noqa: S106 — dummy literal for tests
    redirect_uri="http://localhost:8080/auth/callback",
)


@pytest.fixture
def wired_runtime(monkeypatch: pytest.MonkeyPatch) -> None:
    """Build an OAuthClient with msal mocked, plus a fresh TokenStore."""
    fake_msal_app = MagicMock()
    fake_msal_app.get_authorization_request_url.side_effect = (
        lambda **kwargs: f"https://login.microsoftonline.com/authorize?state={kwargs['state']}"
    )
    monkeypatch.setattr(
        "office_mcp.auth.oauth.msal.ConfidentialClientApplication",
        MagicMock(return_value=fake_msal_app),
    )
    set_oauth_client(OAuthClient(CONFIG))
    set_token_store(TokenStore())


@pytest.mark.asyncio
async def test_wrapper_passes_through_normal_result(
    wired_runtime: None,
) -> None:
    async def ok() -> dict[str, str]:
        return {"summary": "all good"}

    result = await instrumented(ok)()

    assert result == {"summary": "all good"}


@pytest.mark.asyncio
async def test_wrapper_converts_auth_required_to_login_url_response(
    wired_runtime: None,
) -> None:
    async def needs_auth() -> dict[str, str]:
        raise AuthRequiredError("Graph 401")

    with bind_mcp_session_id("session-xyz"):
        result = await instrumented(needs_auth)()

    assert result["auth_required"] is True
    assert result["login_url"].startswith("https://login.microsoftonline.com/")
    assert "Connectez-vous" in result["summary"]


@pytest.mark.asyncio
async def test_wrapper_falls_back_when_pending_store_full(
    monkeypatch: pytest.MonkeyPatch,
    wired_runtime: None,
) -> None:
    """PendingStoreFull must produce an auth_required dict, not a raw 500."""
    from office_mcp.auth.store import TokenStore
    from office_mcp.runtime import get_token_store

    # Force the cap to 0 so any remember_pending raises immediately.
    monkeypatch.setattr(TokenStore, "MAX_PENDING", 0)
    # Re-register the same store under the patched class state.
    get_token_store()  # ensure it exists; cap is class-level so applies live

    async def needs_auth() -> dict[str, str]:
        raise AuthRequiredError("Graph 401")

    with bind_mcp_session_id("session-busy"):
        result = await instrumented(needs_auth)()

    assert result["auth_required"] is True
    assert "login_url" not in result
    assert "Réessayez" in result["summary"]


@pytest.mark.asyncio
async def test_wrapper_falls_back_when_oauth_disabled() -> None:
    """When Azure creds are missing the wrapper must not pretend a
    login URL exists — it returns a clear "feature unavailable" shape."""
    set_oauth_client(None)
    set_token_store(TokenStore())

    async def needs_auth() -> dict[str, str]:
        raise AuthRequiredError("Graph 401")

    result = await instrumented(needs_auth)()

    assert result["auth_required"] is True
    assert "login_url" not in result
    assert "not configured" in result["summary"].lower()


@pytest.mark.asyncio
async def test_wrapper_still_propagates_other_exceptions(
    wired_runtime: None,
) -> None:
    async def boom() -> dict[str, str]:
        raise ValueError("something else broke")

    with pytest.raises(ValueError, match="something else broke"):
        await instrumented(boom)()
