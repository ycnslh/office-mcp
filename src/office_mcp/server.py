"""FastMCP streamable-http server with auth + user context middleware.

Routes:
  GET  /health   — unauthenticated liveness probe (used by Docker, CI, ops).
  ANY  /mcp      — MCP JSON-RPC over streamable HTTP (LibreChat connects here).

Middleware (outermost → innermost):
  BearerAuthMiddleware     — rejects every non-/health request when
                             ``OFFICE_MCP_TOKEN`` is set and the bearer
                             does not match.
  UserContextMiddleware    — extracts ``X-LibreChat-User-Id`` /
                             ``X-LibreChat-User-Email`` into a ContextVar
                             that tools read via :func:`current_user`.

Tools and resources are registered by their respective modules in later
phases. This module only stands up the transport.
"""

from __future__ import annotations

import uvicorn
from fastmcp import FastMCP
from starlette.requests import Request
from starlette.responses import JSONResponse
from starlette.types import ASGIApp, Receive, Scope, Send

from office_mcp import __version__
from office_mcp.config import Settings
from office_mcp.context import UserContext, bind_user

mcp: FastMCP = FastMCP("office-mcp")


# ── HTTP routes ────────────────────────────────────────────────────────


@mcp.custom_route("/health", methods=["GET"])
async def _health(_: Request) -> JSONResponse:
    return JSONResponse(
        {
            "status": "ok",
            "version": __version__,
            "phase": "2-mcp-wired",
            "mcp_ready": True,
        }
    )


# ── ASGI middlewares ───────────────────────────────────────────────────
# Pure ASGI (not BaseHTTPMiddleware) so SSE streams emitted by the MCP
# transport are not buffered into memory.


class BearerAuthMiddleware:
    """Reject requests missing a valid ``Authorization: Bearer <token>``.

    No-ops when ``expected_token`` is ``None``. Always lets ``/health``
    through so Docker healthchecks and uptime probes don't need the token.
    """

    def __init__(self, app: ASGIApp, expected_token: str | None) -> None:
        self.app = app
        self._expected = expected_token

    async def __call__(self, scope: Scope, receive: Receive, send: Send) -> None:
        if scope["type"] != "http" or self._expected is None:
            await self.app(scope, receive, send)
            return
        if scope.get("path") == "/health":
            await self.app(scope, receive, send)
            return
        auth = _header(scope, b"authorization")
        if not auth.startswith("Bearer "):
            await _reject(send, 401, b'{"error":"unauthorized"}')
            return
        if auth.removeprefix("Bearer ").strip() != self._expected:
            await _reject(send, 401, b'{"error":"unauthorized"}')
            return
        await self.app(scope, receive, send)


class UserContextMiddleware:
    """Bind a :class:`UserContext` from LibreChat headers for the request."""

    def __init__(self, app: ASGIApp) -> None:
        self.app = app

    async def __call__(self, scope: Scope, receive: Receive, send: Send) -> None:
        if scope["type"] != "http":
            await self.app(scope, receive, send)
            return
        user_id = _header(scope, b"x-librechat-user-id")
        if not user_id:
            await self.app(scope, receive, send)
            return
        email = _header(scope, b"x-librechat-user-email") or None
        with bind_user(UserContext(user_id=user_id, email=email)):
            await self.app(scope, receive, send)


def _header(scope: Scope, name: bytes) -> str:
    for key, value in scope.get("headers", []):
        if key.lower() == name:
            return value.decode("latin-1")
    return ""


async def _reject(send: Send, status: int, body: bytes) -> None:
    await send(
        {
            "type": "http.response.start",
            "status": status,
            "headers": [(b"content-type", b"application/json")],
        }
    )
    await send({"type": "http.response.body", "body": body})


# ── App factory ────────────────────────────────────────────────────────


def build_app(settings: Settings) -> ASGIApp:
    app = mcp.http_app(path="/mcp")
    app.add_middleware(UserContextMiddleware)
    app.add_middleware(BearerAuthMiddleware, expected_token=settings.token)
    return app


def run() -> None:
    settings = Settings()
    uvicorn.run(
        build_app(settings),
        host=settings.host,
        port=settings.port,
        log_level=settings.log_level.lower(),
    )
