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
            "X-LibreChat-User-Id": "yacine",
            "X-LibreChat-User-Email": "yacine@netwatch.fr",
        },
    )
    assert response.json() == {"user_id": "yacine", "email": "yacine@netwatch.fr"}


@pytest.mark.asyncio
async def test_user_context_absent_when_header_missing() -> None:
    app = _app(Middleware(UserContextMiddleware))
    response = await _get(app, "/whoami")
    assert response.json() == {"user_id": None, "email": None}


@pytest.mark.asyncio
async def test_user_context_email_optional() -> None:
    app = _app(Middleware(UserContextMiddleware))
    response = await _get(
        app, "/whoami", headers={"X-LibreChat-User-Id": "alice"}
    )
    assert response.json() == {"user_id": "alice", "email": None}


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
        app, "/whoami", headers={"X-LibreChat-User-Id": "yacine"}
    )
    assert response.status_code == 401
    assert seen["inner_ran"] is False
