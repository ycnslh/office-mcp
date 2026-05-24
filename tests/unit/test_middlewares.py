"""Unit tests for the bearer-auth and user-context ASGI middlewares."""

from __future__ import annotations

from collections.abc import Callable

import httpx
import pytest
from httpx import ASGITransport
from starlette.applications import Starlette
from starlette.middleware import Middleware
from starlette.requests import Request
from starlette.responses import JSONResponse
from starlette.routing import Route

from office_mcp.context import current_user
from office_mcp.server import BearerAuthMiddleware, UserContextMiddleware


async def _whoami(_: Request) -> JSONResponse:
    user = current_user()
    return JSONResponse(
        {
            "user_id": user.user_id if user else None,
            "email": user.email if user else None,
        }
    )


async def _ok(_: Request) -> JSONResponse:
    return JSONResponse({"ok": True})


def _app(*middlewares: Middleware, routes: list[Route] | None = None) -> Starlette:
    return Starlette(
        routes=routes or [Route("/whoami", _whoami)],
        middleware=list(middlewares),
    )


async def _get(
    app: Starlette, path: str, headers: dict[str, str] | None = None
) -> httpx.Response:
    transport = ASGITransport(app=app)
    async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
        return await client.get(path, headers=headers or {})


# ── UserContextMiddleware ────────────────────────────────────────────


@pytest.mark.asyncio
async def test_user_context_extracted_from_headers() -> None:
    app = _app(Middleware(UserContextMiddleware))
    response = await _get(
        app,
        "/whoami",
        headers={
            "X-LibreChat-User-Id": "alice",
            "X-LibreChat-User-Email": "alice@example.com",
        },
    )
    assert response.json() == {"user_id": "alice", "email": "alice@example.com"}


@pytest.mark.asyncio
async def test_user_context_falls_back_when_header_missing() -> None:
    app = _app(Middleware(UserContextMiddleware, fallback_user_id="demo-user"))
    response = await _get(app, "/whoami")
    assert response.json() == {"user_id": "demo-user", "email": None}


@pytest.mark.asyncio
async def test_user_context_falls_back_on_unsubstituted_placeholder() -> None:
    """LibreChat 0.8.4 forwards {{LIBRECHAT_USER_ID}} literally — treat as missing."""
    app = _app(Middleware(UserContextMiddleware, fallback_user_id="demo-user"))
    response = await _get(
        app,
        "/whoami",
        headers={
            "X-LibreChat-User-Id": "{{LIBRECHAT_USER_ID}}",
            "X-LibreChat-User-Email": "{{LIBRECHAT_USER_EMAIL}}",
        },
    )
    assert response.json() == {"user_id": "demo-user", "email": None}


@pytest.mark.asyncio
async def test_user_context_email_optional() -> None:
    app = _app(Middleware(UserContextMiddleware))
    response = await _get(
        app, "/whoami", headers={"X-LibreChat-User-Id": "alice"}
    )
    assert response.json() == {"user_id": "alice", "email": None}


@pytest.mark.asyncio
async def test_user_context_attaches_session_token_when_available() -> None:
    """Phase 7c: an Mcp-Session-Id header that resolves in the token
    store causes the middleware to stamp the access token onto the
    bound UserContext, so cloud backends can act on the user's behalf.
    """
    import time

    from office_mcp.auth.store import TokenEntry, TokenStore
    from office_mcp.runtime import set_token_store

    store = TokenStore()
    store.put_token(
        "state-X",
        TokenEntry(
            access_token="bound-AT",
            refresh_token=None,
            expires_at=time.time() + 3600,
            user_oid="oid-1",
            user_upn="alice@example.com",
        ),
    )
    store.link_session_to_state("session-S1", "state-X")
    set_token_store(store)

    async def _whoami_with_token(_: Request) -> JSONResponse:
        user = current_user()
        return JSONResponse(
            {"on_behalf_of_token": user.on_behalf_of_token if user else None}
        )

    app = Starlette(
        routes=[Route("/whoami", _whoami_with_token)],
        middleware=[Middleware(UserContextMiddleware)],
    )

    response = await _get(
        app,
        "/whoami",
        headers={
            "X-LibreChat-User-Id": "alice",
            "Mcp-Session-Id": "session-S1",
        },
    )

    assert response.json() == {"on_behalf_of_token": "bound-AT"}


@pytest.mark.asyncio
async def test_user_context_uses_forwarded_oidc_bearer() -> None:
    """LibreChat 0.8.5 forwards the OIDC access token via X-User-Authorization;
    the middleware must use it as on_behalf_of_token without touching the
    legacy session store.
    """

    async def _whoami_with_token(_: Request) -> JSONResponse:
        user = current_user()
        return JSONResponse(
            {"on_behalf_of_token": user.on_behalf_of_token if user else None}
        )

    app = Starlette(
        routes=[Route("/whoami", _whoami_with_token)],
        middleware=[Middleware(UserContextMiddleware)],
    )

    response = await _get(
        app,
        "/whoami",
        headers={
            "X-LibreChat-User-Id": "alice",
            "X-User-Authorization": "Bearer oidc-access-token-123",
        },
    )
    assert response.json() == {"on_behalf_of_token": "oidc-access-token-123"}


@pytest.mark.asyncio
async def test_user_context_ignores_unsubstituted_oidc_placeholder() -> None:
    """When the user is not signed in via OIDC, LibreChat forwards the
    literal ``{{LIBRECHAT_OPENID_ACCESS_TOKEN}}`` placeholder; treat it as
    absent so the legacy session lookup can take over.
    """
    from office_mcp.auth.store import TokenStore
    from office_mcp.runtime import set_token_store

    set_token_store(TokenStore())

    async def _whoami_with_token(_: Request) -> JSONResponse:
        user = current_user()
        return JSONResponse(
            {"on_behalf_of_token": user.on_behalf_of_token if user else None}
        )

    app = Starlette(
        routes=[Route("/whoami", _whoami_with_token)],
        middleware=[Middleware(UserContextMiddleware)],
    )

    response = await _get(
        app,
        "/whoami",
        headers={
            "X-LibreChat-User-Id": "alice",
            "X-User-Authorization": "Bearer {{LIBRECHAT_OPENID_ACCESS_TOKEN}}",
        },
    )
    assert response.json() == {"on_behalf_of_token": None}


@pytest.mark.asyncio
async def test_user_context_oidc_bearer_takes_precedence_over_session() -> None:
    """When both X-User-Authorization and a known Mcp-Session-Id are present,
    the forwarded OIDC token wins (single sign-on path).
    """
    import time

    from office_mcp.auth.store import TokenEntry, TokenStore
    from office_mcp.runtime import set_token_store

    store = TokenStore()
    store.put_token(
        "state-X",
        TokenEntry(
            access_token="legacy-AT",
            refresh_token=None,
            expires_at=time.time() + 3600,
            user_oid="oid-1",
            user_upn="alice@example.com",
        ),
    )
    store.link_session_to_state("session-S1", "state-X")
    set_token_store(store)

    async def _whoami_with_token(_: Request) -> JSONResponse:
        user = current_user()
        return JSONResponse(
            {"on_behalf_of_token": user.on_behalf_of_token if user else None}
        )

    app = Starlette(
        routes=[Route("/whoami", _whoami_with_token)],
        middleware=[Middleware(UserContextMiddleware)],
    )

    response = await _get(
        app,
        "/whoami",
        headers={
            "X-LibreChat-User-Id": "alice",
            "X-User-Authorization": "Bearer oidc-AT",
            "Mcp-Session-Id": "session-S1",
        },
    )
    assert response.json() == {"on_behalf_of_token": "oidc-AT"}


@pytest.mark.asyncio
async def test_user_context_no_token_when_session_unknown() -> None:
    """No leak: an unknown Mcp-Session-Id results in no token attached."""
    from office_mcp.auth.store import TokenStore
    from office_mcp.runtime import set_token_store

    set_token_store(TokenStore())

    async def _whoami_with_token(_: Request) -> JSONResponse:
        user = current_user()
        return JSONResponse(
            {"on_behalf_of_token": user.on_behalf_of_token if user else None}
        )

    app = Starlette(
        routes=[Route("/whoami", _whoami_with_token)],
        middleware=[Middleware(UserContextMiddleware)],
    )

    response = await _get(
        app,
        "/whoami",
        headers={
            "X-LibreChat-User-Id": "alice",
            "Mcp-Session-Id": "session-never-seen",
        },
    )

    assert response.json() == {"on_behalf_of_token": None}


# ── BearerAuthMiddleware ─────────────────────────────────────────────


@pytest.mark.asyncio
async def test_bearer_auth_disabled_when_no_token() -> None:
    app = _app(Middleware(BearerAuthMiddleware, expected_token=None))
    response = await _get(app, "/whoami")
    assert response.status_code == 200


@pytest.mark.asyncio
async def test_bearer_auth_rejects_missing_header() -> None:
    app = _app(Middleware(BearerAuthMiddleware, expected_token="s3cr3t"))
    response = await _get(app, "/whoami")
    assert response.status_code == 401


@pytest.mark.asyncio
async def test_bearer_auth_rejects_wrong_token() -> None:
    app = _app(Middleware(BearerAuthMiddleware, expected_token="s3cr3t"))
    response = await _get(
        app, "/whoami", headers={"Authorization": "Bearer wrong"}
    )
    assert response.status_code == 401


@pytest.mark.asyncio
async def test_bearer_auth_accepts_valid_token() -> None:
    app = _app(Middleware(BearerAuthMiddleware, expected_token="s3cr3t"))
    response = await _get(
        app, "/whoami", headers={"Authorization": "Bearer s3cr3t"}
    )
    assert response.status_code == 200


@pytest.mark.asyncio
async def test_bearer_auth_lets_health_through_without_token() -> None:
    app = _app(
        Middleware(BearerAuthMiddleware, expected_token="s3cr3t"),
        routes=[Route("/health", _ok)],
    )
    response = await _get(app, "/health")
    assert response.status_code == 200


@pytest.mark.asyncio
async def test_bearer_auth_lets_files_endpoint_through_without_token() -> None:
    """`/files/{token}` carries its own auth in the HMAC-signed path token."""
    app = _app(
        Middleware(BearerAuthMiddleware, expected_token="s3cr3t"),
        routes=[Route("/files/{token}", _ok)],
    )
    response = await _get(app, "/files/any-token-value")
    assert response.status_code == 200


# ── Stacking order matches build_app() ───────────────────────────────


@pytest.mark.asyncio
async def test_bearer_short_circuits_before_inner_middlewares() -> None:
    """A 401 from the outer bearer guard must NOT invoke any inner middleware.

    Mirrors the order in :func:`office_mcp.server.build_app`, which calls
    ``add_middleware(UserContext)`` then ``add_middleware(Bearer)`` so that
    Bearer ends up outermost.
    """
    seen: dict[str, bool] = {"inner_ran": False}

    class _Spy:
        def __init__(self, app: Callable) -> None:
            self.app = app

        async def __call__(self, scope, receive, send):  # type: ignore[no-untyped-def]
            seen["inner_ran"] = True
            await self.app(scope, receive, send)

    app = Starlette(routes=[Route("/whoami", _whoami)])
    app.add_middleware(_Spy)  # innermost
    app.add_middleware(BearerAuthMiddleware, expected_token="s3cr3t")  # outermost

    response = await _get(
        app, "/whoami", headers={"X-LibreChat-User-Id": "alice"}
    )
    assert response.status_code == 401
    assert seen["inner_ran"] is False
