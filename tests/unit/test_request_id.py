"""Unit tests for the X-Request-Id correlation middleware."""

from __future__ import annotations

import httpx
import pytest
from httpx import ASGITransport
from starlette.applications import Starlette
from starlette.middleware import Middleware
from starlette.requests import Request
from starlette.responses import JSONResponse
from starlette.routing import Route

from office_mcp.server import RequestIdMiddleware


async def _echo(_: Request) -> JSONResponse:
    import structlog

    bound = structlog.contextvars.get_contextvars()
    return JSONResponse({"request_id_in_context": bound.get("request_id")})


def _app() -> Starlette:
    return Starlette(
        routes=[Route("/x", _echo)],
        middleware=[Middleware(RequestIdMiddleware)],
    )


async def _get(app: Starlette, headers: dict[str, str] | None = None) -> httpx.Response:
    transport = ASGITransport(app=app)
    async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
        return await client.get("/x", headers=headers or {})


@pytest.mark.asyncio
async def test_request_id_generated_when_absent() -> None:
    response = await _get(_app())
    assert response.status_code == 200
    assert response.headers["x-request-id"]
    assert response.json()["request_id_in_context"] == response.headers["x-request-id"]


@pytest.mark.asyncio
async def test_request_id_honours_inbound_header() -> None:
    response = await _get(_app(), headers={"X-Request-Id": "fixed-id-42"})
    assert response.headers["x-request-id"] == "fixed-id-42"
    assert response.json()["request_id_in_context"] == "fixed-id-42"
