"""HTTP-request-driven lookup in :mod:`office_mcp.context`.

The ASGI middleware sets a ContextVar with the per-request user/session
data, but FastMCP runs tool calls inside a long-lived session task whose
context snapshot is taken at session-init time — so a ContextVar set on
a *later* HTTP request does not propagate. ``current_user`` and
``current_mcp_session_id`` therefore prefer FastMCP's per-call
``get_http_request()`` over the ContextVar fallback. These tests pin
that ordering and the fallbacks.
"""

from __future__ import annotations

import time
from collections.abc import Iterator
from unittest.mock import patch

import pytest
from starlette.requests import Request

from office_mcp.auth.store import TokenEntry, TokenStore
from office_mcp.context import (
    UserContext,
    bind_mcp_session_id,
    bind_user,
    current_mcp_session_id,
    current_user,
)
from office_mcp.runtime import set_token_store


def _request(headers: dict[str, str]) -> Request:
    """Build a minimal Starlette Request with the given headers."""
    raw = [(k.lower().encode(), v.encode()) for k, v in headers.items()]
    scope = {
        "type": "http",
        "method": "POST",
        "path": "/mcp",
        "headers": raw,
        "query_string": b"",
    }
    return Request(scope)


@pytest.fixture
def isolated_token_store() -> Iterator[TokenStore]:
    """Per-test TokenStore so cross-test state can't leak through runtime."""
    store = TokenStore()
    set_token_store(store)
    try:
        yield store
    finally:
        set_token_store(None)


def test_current_user_reads_from_http_request_first() -> None:
    """``get_http_request()`` wins over a stale ContextVar binding."""
    fake_request = _request(
        {"X-LibreChat-User-Id": "alice", "X-LibreChat-User-Email": "alice@example.com"}
    )
    stale = UserContext(user_id="frozen-from-session-init")
    with bind_user(stale), patch(
        "fastmcp.server.dependencies.get_http_request", return_value=fake_request
    ):
        ctx = current_user()
    assert ctx is not None
    assert ctx.user_id == "alice"
    assert ctx.email == "alice@example.com"


def test_current_user_falls_back_to_context_var_outside_http() -> None:
    """No active HTTP request → ContextVar (used by tests, /files/{token})."""
    bound = UserContext(user_id="bound-user", email="b@x.fr")
    with bind_user(bound):
        ctx = current_user()
    assert ctx == bound


def test_current_user_returns_none_when_nothing_set() -> None:
    assert current_user() is None


def test_current_user_falls_back_to_default_for_placeholder_header() -> None:
    """LibreChat 0.8.4 forwards literal ``{{LIBRECHAT_USER_ID}}`` — treat as missing."""
    fake_request = _request({"X-LibreChat-User-Id": "{{LIBRECHAT_USER_ID}}"})
    with patch(
        "fastmcp.server.dependencies.get_http_request", return_value=fake_request
    ):
        ctx = current_user()
    assert ctx is not None
    assert ctx.user_id == "demo-user"
    assert ctx.email is None


def test_current_user_attaches_token_from_session_index(
    isolated_token_store: TokenStore,
) -> None:
    """Live token store lookup populates ``on_behalf_of_token`` per call."""
    isolated_token_store.put_token(
        "state-X",
        TokenEntry(
            access_token="bound-AT",
            refresh_token=None,
            expires_at=time.time() + 3600,
            user_oid="oid-1",
            user_upn="alice@example.com",
        ),
    )
    isolated_token_store.link_session_to_state("session-S1", "state-X")

    fake_request = _request(
        {"X-LibreChat-User-Id": "alice", "Mcp-Session-Id": "session-S1"}
    )
    with patch(
        "fastmcp.server.dependencies.get_http_request", return_value=fake_request
    ):
        ctx = current_user()
    assert ctx is not None
    assert ctx.on_behalf_of_token == "bound-AT"


def test_current_user_uses_forwarded_oidc_bearer(
    isolated_token_store: TokenStore,
) -> None:
    """LibreChat 0.8.5 forwards the OIDC access token via X-User-Authorization."""
    fake_request = _request(
        {
            "X-LibreChat-User-Id": "alice",
            "X-User-Authorization": "Bearer oidc-access-token-xyz",
        }
    )
    with patch(
        "fastmcp.server.dependencies.get_http_request", return_value=fake_request
    ):
        ctx = current_user()
    assert ctx is not None
    assert ctx.on_behalf_of_token == "oidc-access-token-xyz"


def test_current_user_oidc_bearer_takes_precedence(
    isolated_token_store: TokenStore,
) -> None:
    """Forwarded OIDC token wins over the legacy session-store token."""
    isolated_token_store.put_token(
        "state-X",
        TokenEntry(
            access_token="legacy-AT",
            refresh_token=None,
            expires_at=time.time() + 3600,
            user_oid="oid-1",
            user_upn="alice@example.com",
        ),
    )
    isolated_token_store.link_session_to_state("session-S1", "state-X")

    fake_request = _request(
        {
            "X-LibreChat-User-Id": "alice",
            "Mcp-Session-Id": "session-S1",
            "X-User-Authorization": "Bearer oidc-AT",
        }
    )
    with patch(
        "fastmcp.server.dependencies.get_http_request", return_value=fake_request
    ):
        ctx = current_user()
    assert ctx is not None
    assert ctx.on_behalf_of_token == "oidc-AT"


def test_current_user_ignores_unsubstituted_oidc_placeholder(
    isolated_token_store: TokenStore,
) -> None:
    """Literal ``{{LIBRECHAT_OPENID_ACCESS_TOKEN}}`` → treat as absent, fallback."""
    fake_request = _request(
        {
            "X-LibreChat-User-Id": "alice",
            "X-User-Authorization": "Bearer {{LIBRECHAT_OPENID_ACCESS_TOKEN}}",
        }
    )
    with patch(
        "fastmcp.server.dependencies.get_http_request", return_value=fake_request
    ):
        ctx = current_user()
    assert ctx is not None
    assert ctx.on_behalf_of_token is None


def test_current_user_no_token_when_session_unknown(
    isolated_token_store: TokenStore,
) -> None:
    fake_request = _request(
        {"X-LibreChat-User-Id": "alice", "Mcp-Session-Id": "session-never-seen"}
    )
    with patch(
        "fastmcp.server.dependencies.get_http_request", return_value=fake_request
    ):
        ctx = current_user()
    assert ctx is not None
    assert ctx.on_behalf_of_token is None


def test_current_mcp_session_id_reads_from_http_request_first() -> None:
    """The session id is the load-bearing key — pin its lookup explicitly."""
    fake_request = _request(
        {"X-LibreChat-User-Id": "alice", "Mcp-Session-Id": "session-from-http"}
    )
    with bind_mcp_session_id("stale-from-init"), patch(
        "fastmcp.server.dependencies.get_http_request", return_value=fake_request
    ):
        assert current_mcp_session_id() == "session-from-http"


def test_current_mcp_session_id_falls_back_to_context_var() -> None:
    with bind_mcp_session_id("ctx-var-session"):
        assert current_mcp_session_id() == "ctx-var-session"


def test_current_mcp_session_id_none_outside_request() -> None:
    assert current_mcp_session_id() is None


def test_current_mcp_session_id_none_when_header_missing_in_http() -> None:
    fake_request = _request({"X-LibreChat-User-Id": "alice"})
    with patch(
        "fastmcp.server.dependencies.get_http_request", return_value=fake_request
    ):
        assert current_mcp_session_id() is None


def test_get_http_request_runtime_error_falls_through_to_context_var() -> None:
    """A FastMCP RuntimeError (no active request) falls back gracefully."""
    bound = UserContext(user_id="from-ctxvar")
    with bind_user(bound), patch(
        "fastmcp.server.dependencies.get_http_request",
        side_effect=RuntimeError("No active HTTP request found."),
    ):
        ctx = current_user()
    assert ctx == bound
