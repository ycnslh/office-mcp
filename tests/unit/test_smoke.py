"""Phase 0 smoke tests — package imports and the /health endpoint responds.

Real test suites land alongside their features in Phases 1+.
"""

from __future__ import annotations

from starlette.testclient import TestClient


def test_package_imports() -> None:
    import office_mcp

    assert office_mcp.__version__


def test_health_endpoint() -> None:
    from office_mcp.server import build_app

    with TestClient(build_app()) as client:
        response = client.get("/health")

    assert response.status_code == 200
    payload = response.json()
    assert payload["status"] == "ok"
    assert payload["mcp_ready"] is False
    assert "version" in payload
