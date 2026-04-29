"""End-to-end test: a real MCP client speaks streamable-http to our server.

We run the full app under uvicorn on an ephemeral port (so lifespan events
fire and FastMCP's session manager initializes properly), then connect with
the official MCP SDK client and complete an ``initialize`` + ``list_tools``
round-trip.
"""

from __future__ import annotations

import socket
import threading
from collections.abc import Iterator
from typing import Any

import httpx
import pytest
import uvicorn
from mcp import ClientSession
from mcp.client.streamable_http import streamablehttp_client

from office_mcp.config import Settings
from office_mcp.server import build_app


def _free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


@pytest.fixture
def running_server(tmp_path) -> Iterator[str]:  # type: ignore[no-untyped-def]
    port = _free_port()
    settings = Settings(
        host="127.0.0.1",
        port=port,
        token=None,
        log_level="warning",
        local_root=tmp_path,
    )
    config = uvicorn.Config(
        build_app(settings),
        host="127.0.0.1",
        port=port,
        log_level="warning",
        access_log=False,
    )
    server = uvicorn.Server(config)
    thread = threading.Thread(target=server.run, daemon=True)
    thread.start()

    base_url = f"http://127.0.0.1:{port}"
    deadline_attempts = 50
    for _ in range(deadline_attempts):
        try:
            response = httpx.get(f"{base_url}/health", timeout=0.5)
            if response.status_code == 200:
                break
        except httpx.HTTPError:
            pass
        threading.Event().wait(0.1)
    else:
        server.should_exit = True
        thread.join(timeout=5)
        raise RuntimeError("uvicorn did not become ready in 5s")

    try:
        yield base_url
    finally:
        server.should_exit = True
        thread.join(timeout=5)


@pytest.mark.asyncio
async def test_mcp_initialize_and_list_tools(running_server: str) -> None:
    async with (
        streamablehttp_client(f"{running_server}/mcp") as (read, write, _),
        ClientSession(read, write) as session,
    ):
        init: Any = await session.initialize()
        assert init.serverInfo.name == "office-mcp"

        tools = await session.list_tools()
        tool_names = {t.name for t in tools.tools}
        # Phases 3-6: docx + xlsx + pptx + fs + docx_to_pdf.
        assert {
            "docx_create",
            "docx_read",
            "docx_append_section",
            "docx_insert_table",
            "docx_to_pdf",
            "xlsx_create",
            "xlsx_read_range",
            "xlsx_write_range",
            "xlsx_apply_formula",
            "pptx_create",
            "pptx_add_title_slide",
            "pptx_add_content_slide",
            "pptx_add_two_content_slide",
            "pptx_add_section_slide",
            "pptx_set_slide_title",
            "pptx_set_slide_bullets",
            "pptx_read",
            "fs_list",
            "fs_stat",
            "fs_copy",
            "fs_delete",
        }.issubset(tool_names)

        resources = await session.list_resources()
        resource_uris = {str(r.uri) for r in resources.resources}
        assert "office://rules/docx" in resource_uris
        assert "office://rules/xlsx" in resource_uris
        assert "office://rules/pptx" in resource_uris

    # /metrics must be reachable on the same uvicorn without a bearer.
    async with httpx.AsyncClient() as client:
        metrics_resp = await client.get(f"{running_server}/metrics", timeout=5)
    assert metrics_resp.status_code == 200
    assert "office_mcp_tool_calls_total" in metrics_resp.text
