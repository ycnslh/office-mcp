"""End-to-end test of ``/auth/login`` → ``/auth/callback`` → ``/auth/status``.

A real uvicorn is started so middleware ordering, route registration
and redirect handling match production. MSAL is patched at the module
level — Azure AD is never contacted.
"""

from __future__ import annotations

import socket
import threading
import time
from collections.abc import Iterator
from pathlib import Path
from unittest.mock import MagicMock
from urllib.parse import parse_qs, urlparse

import httpx
import pytest
import uvicorn

from office_mcp.config import Settings
from office_mcp.server import build_app


def _free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def _wait_ready(base_url: str, timeout_s: float = 5.0) -> None:
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        try:
            if httpx.get(f"{base_url}/health", timeout=0.5).status_code == 200:
                return
        except httpx.HTTPError:
            pass
        time.sleep(0.1)
    raise RuntimeError(f"uvicorn did not become ready in {timeout_s}s")


def _state_from_redirect(location: str) -> str:
    """Extract the ``state`` query param from a Location header."""
    qs = parse_qs(urlparse(location).query)
    return qs["state"][0]


@pytest.fixture
def fake_msal(monkeypatch: pytest.MonkeyPatch) -> MagicMock:
    fake_instance = MagicMock()
    fake_instance.get_authorization_request_url.side_effect = lambda **kwargs: (
        f"https://login.microsoftonline.com/authorize?state={kwargs['state']}"
    )
    monkeypatch.setattr(
        "office_mcp.auth.oauth.msal.ConfidentialClientApplication",
        MagicMock(return_value=fake_instance),
    )
    return fake_instance


def _start_server(settings: Settings) -> tuple[uvicorn.Server, threading.Thread, str]:
    config = uvicorn.Config(
        build_app(settings),
        host="127.0.0.1",
        port=settings.port,
        log_level="warning",
        access_log=False,
    )
    server = uvicorn.Server(config)
    thread = threading.Thread(target=server.run, daemon=True)
    thread.start()
    base_url = f"http://127.0.0.1:{settings.port}"
    _wait_ready(base_url)
    return server, thread, base_url


@pytest.fixture
def running_server(
    tmp_path: Path, fake_msal: MagicMock
) -> Iterator[tuple[str, MagicMock]]:
    port = _free_port()
    settings = Settings(
        host="127.0.0.1",
        port=port,
        token=None,
        log_level="warning",
        local_root=tmp_path,
        public_base_url=f"http://127.0.0.1:{port}",
        azure_tenant_id="tenant-id",
        azure_client_id="client-id",
        azure_client_secret="secret",  # noqa: S106 — dummy literal for tests
    )
    server, thread, base_url = _start_server(settings)
    try:
        yield base_url, fake_msal
    finally:
        server.should_exit = True
        thread.join(timeout=5)


def test_login_redirects_to_microsoft_with_server_generated_state(
    running_server: tuple[str, MagicMock],
) -> None:
    base, _ = running_server

    response = httpx.get(f"{base}/auth/login", follow_redirects=False, timeout=5)

    assert response.status_code == 302
    location = response.headers["location"]
    assert "login.microsoftonline.com" in location
    state = _state_from_redirect(location)
    # secrets.token_urlsafe(32) yields ≥ 43 url-safe characters.
    assert len(state) >= 32


def test_login_ignores_caller_supplied_state(
    running_server: tuple[str, MagicMock],
) -> None:
    """A weak/known state from a caller must NEVER be honoured.

    The state is a capability — it is generated server-side only.
    """
    base, _ = running_server

    response = httpx.get(
        f"{base}/auth/login?state=weak-attacker-supplied",
        follow_redirects=False,
        timeout=5,
    )

    assert response.status_code == 302
    state = _state_from_redirect(response.headers["location"])
    assert state != "weak-attacker-supplied"
    assert len(state) >= 32


def test_callback_exchanges_code_and_marks_authenticated(
    running_server: tuple[str, MagicMock],
) -> None:
    base, msal_app = running_server
    msal_app.acquire_token_by_authorization_code.return_value = {
        "access_token": "AT",
        "refresh_token": "RT",
        "expires_in": 3600,
        "id_token_claims": {
            "oid": "oid-1",
            "preferred_username": "alice@contoso.com",
            "email": "alice@contoso.com",
        },
    }

    # Drive /auth/login to register a fresh server-generated state.
    login = httpx.get(f"{base}/auth/login", follow_redirects=False, timeout=5)
    state = _state_from_redirect(login.headers["location"])

    # Microsoft would redirect back here.
    callback = httpx.get(
        f"{base}/auth/callback?code=auth-code&state={state}", timeout=5
    )
    assert callback.status_code == 200
    # The HTML page intentionally exposes the UPN to the user's browser.
    assert "alice@contoso.com" in callback.text

    # /auth/status must NOT echo identity claims — only liveness.
    status = httpx.get(f"{base}/auth/status?state={state}", timeout=5)
    assert status.status_code == 200
    assert status.json() == {"authenticated": True}


def test_callback_rejects_unknown_state(running_server: tuple[str, MagicMock]) -> None:
    base, _ = running_server

    response = httpx.get(
        f"{base}/auth/callback?code=anything&state=never-issued", timeout=5
    )

    assert response.status_code == 400
    assert "unknown" in response.json()["error"].lower()


def test_callback_propagates_microsoft_error(
    running_server: tuple[str, MagicMock],
) -> None:
    base, _ = running_server

    response = httpx.get(
        f"{base}/auth/callback?error=access_denied&error_description=user+canceled",
        timeout=5,
    )

    assert response.status_code == 400
    assert "access_denied" in response.text


def test_callback_returns_400_on_missing_code(
    running_server: tuple[str, MagicMock],
) -> None:
    base, _ = running_server

    response = httpx.get(f"{base}/auth/callback?state=foo", timeout=5)

    assert response.status_code == 400
    assert "missing" in response.json()["error"].lower()


def test_status_unauthenticated_when_state_missing(
    running_server: tuple[str, MagicMock],
) -> None:
    base, _ = running_server

    response = httpx.get(f"{base}/auth/status?state=never-set", timeout=5)

    assert response.status_code == 200
    assert response.json() == {"authenticated": False}


def test_status_400_when_state_query_param_omitted(
    running_server: tuple[str, MagicMock],
) -> None:
    base, _ = running_server

    response = httpx.get(f"{base}/auth/status", timeout=5)

    assert response.status_code == 400


def test_auth_endpoints_exempt_from_bearer_when_token_set(
    tmp_path: Path, fake_msal: MagicMock
) -> None:
    """A bearer-protected server still serves /auth/login without the token."""
    port = _free_port()
    settings = Settings(
        host="127.0.0.1",
        port=port,
        token="shared-bearer-secret",  # noqa: S106 — dummy literal for tests
        log_level="warning",
        local_root=tmp_path,
        public_base_url=f"http://127.0.0.1:{port}",
        azure_tenant_id="tenant-id",
        azure_client_id="client-id",
        azure_client_secret="secret",  # noqa: S106 — dummy literal for tests
    )
    server, thread, base = _start_server(settings)
    try:
        response = httpx.get(
            f"{base}/auth/login", follow_redirects=False, timeout=5
        )
        assert response.status_code == 302
    finally:
        server.should_exit = True
        thread.join(timeout=5)


def test_auth_endpoints_return_503_when_oauth_disabled(tmp_path: Path) -> None:
    """No Azure creds → /auth/* answers 503 instead of crashing."""
    port = _free_port()
    settings = Settings(
        host="127.0.0.1",
        port=port,
        token=None,
        log_level="warning",
        local_root=tmp_path,
        # No public_base_url, no Azure creds.
    )
    server, thread, base = _start_server(settings)
    try:
        response = httpx.get(f"{base}/auth/status?state=foo", timeout=5)
        assert response.status_code == 503
    finally:
        server.should_exit = True
        thread.join(timeout=5)


def test_auth_start_redirects_to_microsoft_using_pending_state(
    running_server: tuple[str, MagicMock],
) -> None:
    """``/auth/start/{state}`` is the short, our-domain redirect that
    tools surface to the LLM. It must look up the pending claim and
    302-redirect to the upstream Microsoft authorize URL with the same
    state — without consuming the claim, so the callback can still
    redeem it.
    """
    from office_mcp.auth.flow import start_oauth_flow

    base, _ = running_server

    # A tool minted a state via auth_required_response → start_oauth_flow.
    state, _ = start_oauth_flow(mcp_session_id="session-A")

    response = httpx.get(
        f"{base}/auth/start/{state}", follow_redirects=False, timeout=5
    )

    assert response.status_code == 302
    location = response.headers["location"]
    assert "login.microsoftonline.com" in location
    assert _state_from_redirect(location) == state


def test_auth_start_404s_on_unknown_state(
    running_server: tuple[str, MagicMock],
) -> None:
    base, _ = running_server

    response = httpx.get(f"{base}/auth/start/never-issued", timeout=5)

    assert response.status_code == 404


def test_auth_start_does_not_consume_pending_claim(
    running_server: tuple[str, MagicMock],
) -> None:
    """If /auth/start redeemed the claim, the callback below would 400
    on the same state — that would break the whole flow."""
    from office_mcp.auth.flow import start_oauth_flow
    from office_mcp.runtime import get_token_store

    base, msal_app = running_server
    msal_app.acquire_token_by_authorization_code.return_value = {
        "access_token": "AT",
        "refresh_token": "RT",
        "expires_in": 3600,
        "id_token_claims": {
            "oid": "oid-1",
            "preferred_username": "alice@contoso.com",
        },
    }
    state, _ = start_oauth_flow(mcp_session_id="session-peek")

    # User clicks the short URL (peek path).
    httpx.get(f"{base}/auth/start/{state}", follow_redirects=False, timeout=5)
    # Microsoft redirects back. The callback must still find the
    # pending claim — i.e. /auth/start did not pop it.
    callback = httpx.get(
        f"{base}/auth/callback?code=auth-code&state={state}", timeout=5
    )
    assert callback.status_code == 200
    assert get_token_store().get_token_for_session("session-peek") is not None


def test_auth_start_exempt_from_bearer_when_token_set(
    tmp_path: Path, fake_msal: MagicMock
) -> None:
    """Tools call /auth/start before the user has any bearer; like
    /auth/login it must work without ``Authorization``."""
    from office_mcp.auth.flow import start_oauth_flow

    port = _free_port()
    settings = Settings(
        host="127.0.0.1",
        port=port,
        token="shared-bearer-secret",  # noqa: S106 — dummy literal for tests
        log_level="warning",
        local_root=tmp_path,
        public_base_url=f"http://127.0.0.1:{port}",
        azure_tenant_id="tenant-id",
        azure_client_id="client-id",
        azure_client_secret="secret",  # noqa: S106 — dummy literal for tests
    )
    server, thread, base = _start_server(settings)
    try:
        state, _ = start_oauth_flow(mcp_session_id="session-bearer")
        response = httpx.get(
            f"{base}/auth/start/{state}", follow_redirects=False, timeout=5
        )
        assert response.status_code == 302
    finally:
        server.should_exit = True
        thread.join(timeout=5)


def test_callback_links_session_when_started_by_a_tool(
    running_server: tuple[str, MagicMock],
) -> None:
    """Phase 7c integration: a tool calls ``start_oauth_flow(session_id)``,
    the user completes consent on Microsoft, the callback hits us — and
    the token must now be retrievable by ``mcp_session_id`` so the next
    tool call in the same chat session resolves it without re-auth."""
    from office_mcp.auth.flow import start_oauth_flow
    from office_mcp.runtime import get_token_store

    base, msal_app = running_server
    msal_app.acquire_token_by_authorization_code.return_value = {
        "access_token": "delegated-AT",
        "refresh_token": "RT",
        "expires_in": 3600,
        "id_token_claims": {
            "oid": "oid-alice",
            "preferred_username": "alice@example.com",
            "email": "alice@example.com",
        },
    }

    # Tool side: mint a state bound to "session-S1".
    # NOTE: in production a tool runs inside the uvicorn worker, but
    # here we mutate the same in-process TokenStore directly. The test
    # thread and the server thread share the singleton wired by
    # build_app() — that's the only reason this works without a real
    # tool call routed through /mcp.
    state, _ = start_oauth_flow(mcp_session_id="session-S1")

    # Microsoft side: redirect back to /auth/callback with that state.
    callback = httpx.get(
        f"{base}/auth/callback?code=auth-code&state={state}", timeout=5
    )
    assert callback.status_code == 200

    # Tool side again, in the same MCP transport session.
    entry = get_token_store().get_token_for_session("session-S1")
    assert entry is not None
    assert entry.access_token == "delegated-AT"
    assert entry.user_upn == "alice@example.com"
