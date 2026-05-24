"""Unit tests for the per-user token-bucket rate limiter."""

from __future__ import annotations

import httpx
import pytest
from httpx import ASGITransport
from starlette.applications import Starlette
from starlette.middleware import Middleware
from starlette.requests import Request
from starlette.responses import JSONResponse
from starlette.routing import Route

from office_mcp.context import UserContext, bind_user
from office_mcp.server import RateLimitMiddleware, _TokenBucket


def test_token_bucket_starts_full() -> None:
    bucket = _TokenBucket(capacity=5)
    assert all(bucket.take() for _ in range(5))
    assert bucket.take() is False


def test_token_bucket_refills_over_time() -> None:
    bucket = _TokenBucket(capacity=60)  # 1 token / second
    assert all(bucket.take() for _ in range(60))
    assert bucket.take() is False
    # rewind last_refill to simulate elapsed time
    bucket.last_refill -= 2.0  # two seconds ago
    assert bucket.take() is True
    assert bucket.take() is True
    assert bucket.take() is False


def test_retry_after_is_at_least_one_second() -> None:
    bucket = _TokenBucket(capacity=60)
    for _ in range(60):
        bucket.take()
    assert bucket.retry_after_seconds() >= 1


# ── middleware-level tests ──────────────────────────────────────────


async def _ok(_: Request) -> JSONResponse:
    return JSONResponse({"ok": True})


def _app(capacity: int) -> Starlette:
    return Starlette(
        routes=[Route("/mcp/x", _ok), Route("/health", _ok)],
        middleware=[Middleware(RateLimitMiddleware, requests_per_minute=capacity)],
    )


async def _get(app: Starlette, path: str, user_id: str | None = "alice") -> httpx.Response:
    transport = ASGITransport(app=app)
    async with httpx.AsyncClient(transport=transport, base_url="http://test") as client:
        if user_id:
            with bind_user(UserContext(user_id=user_id)):
                return await client.get(path)
        return await client.get(path)


@pytest.mark.asyncio
async def test_rate_limit_allows_under_capacity() -> None:
    app = _app(capacity=3)
    for _ in range(3):
        response = await _get(app, "/mcp/x")
        assert response.status_code == 200


@pytest.mark.asyncio
async def test_rate_limit_blocks_over_capacity_with_429() -> None:
    app = _app(capacity=2)
    for _ in range(2):
        assert (await _get(app, "/mcp/x")).status_code == 200
    response = await _get(app, "/mcp/x")
    assert response.status_code == 429
    assert response.headers["retry-after"]
    assert response.json()["error"] == "rate limit exceeded"


@pytest.mark.asyncio
async def test_rate_limit_per_user_isolation() -> None:
    app = _app(capacity=1)
    # alice burns her token
    assert (await _get(app, "/mcp/x", user_id="alice")).status_code == 200
    assert (await _get(app, "/mcp/x", user_id="alice")).status_code == 429
    # bob is unaffected
    assert (await _get(app, "/mcp/x", user_id="bob")).status_code == 200


@pytest.mark.asyncio
async def test_rate_limit_skips_non_mcp_paths() -> None:
    app = _app(capacity=1)
    assert (await _get(app, "/mcp/x")).status_code == 200
    # /health unrelated to /mcp — must not be limited
    for _ in range(5):
        assert (await _get(app, "/health")).status_code == 200


@pytest.mark.asyncio
async def test_rate_limit_disabled_when_capacity_is_zero() -> None:
    app = _app(capacity=0)
    for _ in range(10):
        assert (await _get(app, "/mcp/x")).status_code == 200
