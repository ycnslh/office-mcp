"""Smoke tests — package import + /health endpoint shape."""

from __future__ import annotations

from starlette.testclient import TestClient

from office_mcp.config import Settings
from office_mcp.server import build_app


def test_package_imports() -> None:
    import office_mcp

    assert office_mcp.__version__


def test_health_endpoint() -> None:
    with TestClient(build_app(Settings(token=None))) as client:
        response = client.get("/health")

    assert response.status_code == 200
    payload = response.json()
    assert payload["status"] == "ok"
    assert payload["mcp_ready"] is True
    assert payload["phase"].startswith("3-")
    assert "version" in payload


def test_health_reachable_even_with_bearer_configured() -> None:
    """The healthcheck must NEVER require a token (Docker probes don't have one)."""
    with TestClient(build_app(Settings(token="some-token"))) as client:
        response = client.get("/health")

    assert response.status_code == 200
