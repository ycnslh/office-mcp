"""Integration test: the ``/files/{token}`` download endpoint.

We stand up a real uvicorn on an ephemeral port so the Starlette
lifespan (and therefore FastMCP's session manager) initialises the same
way it does in production. Pure ASGITransport would work for plain
routes but would not exercise the rest of the stack.
"""

from __future__ import annotations

import socket
import threading
import time
from collections.abc import Iterator
from pathlib import Path

import httpx
import pytest
import uvicorn

from office_mcp.config import Settings
from office_mcp.context import UserContext, bind_user
from office_mcp.runtime import build_download_url, set_registry
from office_mcp.server import build_app
from office_mcp.storage import LocalBackend, StorageRegistry
from office_mcp.utils import signing


def _free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


@pytest.fixture
def running_server(tmp_path: Path) -> Iterator[tuple[str, Settings]]:
    port = _free_port()
    settings = Settings(
        host="127.0.0.1",
        port=port,
        token=None,
        log_level="warning",
        local_root=tmp_path,
        signing_secret="unit-test-secret",
        download_base_url=f"http://127.0.0.1:{port}",
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
    for _ in range(50):
        try:
            if httpx.get(f"http://127.0.0.1:{port}/health", timeout=0.5).status_code == 200:
                break
        except httpx.HTTPError:
            pass
        time.sleep(0.1)
    else:
        server.should_exit = True
        thread.join(timeout=5)
        raise RuntimeError("uvicorn did not become ready in 5s")
    try:
        yield f"http://127.0.0.1:{port}", settings
    finally:
        server.should_exit = True
        thread.join(timeout=5)


def _seed_file(settings: Settings, user_id: str, uri: str, data: bytes) -> None:
    import asyncio

    registry = StorageRegistry()
    registry.register(LocalBackend(root=settings.local_root))
    set_registry(registry)
    with bind_user(UserContext(user_id=user_id)):
        asyncio.run(
            registry.get(uri).write(
                uri, data, user_ctx=UserContext(user_id=user_id)
            )
        )


def test_download_serves_file_with_valid_token(
    running_server: tuple[str, Settings],
) -> None:
    base_url, settings = running_server
    _seed_file(settings, user_id="yacine", uri="local://x.txt", data=b"hello world")

    with bind_user(UserContext(user_id="yacine")):
        download_url = build_download_url("local://x.txt", "yacine")

    assert download_url is not None
    response = httpx.get(download_url, timeout=5)
    assert response.status_code == 200
    assert response.content == b"hello world"
    assert "attachment" in response.headers["content-disposition"]
    assert "x.txt" in response.headers["content-disposition"]


def test_download_rejects_tampered_token(
    running_server: tuple[str, Settings],
) -> None:
    base_url, settings = running_server
    _seed_file(settings, user_id="alice", uri="local://secret.txt", data=b"top secret")

    # Valid token for alice — then tamper it
    with bind_user(UserContext(user_id="alice")):
        url = build_download_url("local://secret.txt", "alice")
    assert url is not None
    bad_url = url + "corrupt"

    response = httpx.get(bad_url, timeout=5)
    assert response.status_code == 403


def test_download_rejects_expired_token(
    running_server: tuple[str, Settings],
) -> None:
    base_url, settings = running_server
    _seed_file(settings, user_id="bob", uri="local://b.txt", data=b"b")

    # Sign with negative TTL so the token is already expired
    expired = signing.sign(
        user_id="bob",
        uri="local://b.txt",
        secret=settings.signing_secret or "",
        ttl_seconds=-1,
    )
    response = httpx.get(f"{base_url}/files/{expired}", timeout=5)
    assert response.status_code == 403


def test_download_returns_404_for_missing_file(
    running_server: tuple[str, Settings],
) -> None:
    base_url, settings = running_server
    # no file seeded — just sign a URL that points at nothing
    token = signing.sign(
        user_id="ghost",
        uri="local://nope.txt",
        secret=settings.signing_secret or "",
    )
    response = httpx.get(f"{base_url}/files/{token}", timeout=5)
    assert response.status_code == 404
