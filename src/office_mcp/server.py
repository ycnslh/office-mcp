"""HTTP entry point.

Phase 0 ships a tiny Starlette app exposing only ``/health`` so the Docker
container starts and CI/CD healthchecks pass. The full FastMCP application
(``/mcp`` streamable-http endpoint, tool/resource registration, user-context
middleware) is wired in Phase 2 — see CLAUDE.md §8.
"""

from __future__ import annotations

import os

import uvicorn
from starlette.applications import Starlette
from starlette.requests import Request
from starlette.responses import JSONResponse
from starlette.routing import Route

from office_mcp import __version__


async def _health(_: Request) -> JSONResponse:
    return JSONResponse(
        {
            "status": "ok",
            "version": __version__,
            "phase": "0-bootstrap",
            "mcp_ready": False,
        }
    )


def build_app() -> Starlette:
    return Starlette(routes=[Route("/health", _health, methods=["GET"])])


def run() -> None:
    host = os.environ.get("OFFICE_MCP_HOST", "0.0.0.0")  # noqa: S104
    port = int(os.environ.get("OFFICE_MCP_PORT", "8080"))
    uvicorn.run(build_app(), host=host, port=port, log_level="info")
