"""Integration tests for the /metrics endpoint and the enriched /health."""

from __future__ import annotations

from pathlib import Path

import pytest
from starlette.testclient import TestClient

from office_mcp.config import Settings
from office_mcp.context import UserContext, bind_user
from office_mcp.server import build_app


def test_health_lists_registered_backend_status(tmp_path: Path) -> None:
    with TestClient(build_app(Settings(local_root=tmp_path))) as client:
        response = client.get("/health")
    assert response.status_code == 200
    payload = response.json()
    assert payload["status"] == "ok"
    assert payload["backends"]["local"] == "ok"


def test_metrics_endpoint_exposes_prometheus_text(tmp_path: Path) -> None:
    with TestClient(build_app(Settings(local_root=tmp_path))) as client:
        response = client.get("/metrics")
    assert response.status_code == 200
    assert response.headers["content-type"].startswith("text/plain")
    body = response.text
    assert "office_mcp_tool_calls_total" in body
    assert "office_mcp_tool_duration_seconds" in body
    assert "office_mcp_downloads_total" in body
    assert "office_mcp_rate_limited_total" in body


def test_metrics_endpoint_bypasses_bearer(tmp_path: Path) -> None:
    settings = Settings(local_root=tmp_path, token="some-bearer")
    with TestClient(build_app(settings)) as client:
        response = client.get("/metrics")
    assert response.status_code == 200


@pytest.mark.asyncio
async def test_instrumented_wrapper_increments_counters() -> None:
    from prometheus_client import generate_latest

    from office_mcp import metrics
    from office_mcp.instrumentation import instrumented

    async def synthetic_tool(value: int) -> int:
        return value * 2

    wrapped = instrumented(synthetic_tool)
    with bind_user(UserContext(user_id="alice")):
        assert await wrapped(21) == 42

    body = generate_latest(metrics.REGISTRY).decode()
    assert 'tool="synthetic_tool"' in body
    assert 'status="ok"' in body
    # histogram bucket lines for the same tool
    assert 'office_mcp_tool_duration_seconds_bucket{le="0.05",tool="synthetic_tool"}' in body


@pytest.mark.asyncio
async def test_instrumented_wrapper_records_failures() -> None:
    from prometheus_client import generate_latest

    from office_mcp import metrics
    from office_mcp.instrumentation import instrumented

    async def boom() -> None:
        raise RuntimeError("kaboom")

    wrapped = instrumented(boom)
    with (
        bind_user(UserContext(user_id="bob")),
        pytest.raises(RuntimeError, match="kaboom"),
    ):
        await wrapped()

    body = generate_latest(metrics.REGISTRY).decode()
    assert 'office_mcp_tool_calls_total{status="error",tool="boom"}' in body
