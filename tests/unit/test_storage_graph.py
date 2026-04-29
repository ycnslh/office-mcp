"""Unit tests for the Microsoft Graph (OneDrive) storage backend.

We use ``pytest-httpx`` to stub Graph responses so the suite never
contacts ``graph.microsoft.com``. The tests focus on the URL shape we
emit, the headers we attach, and the way we map provider statuses to
office-mcp errors.
"""

from __future__ import annotations

from collections.abc import AsyncIterator
from pathlib import Path

import pytest
import pytest_asyncio
from pytest_httpx import HTTPXMock

from office_mcp.context import UserContext
from office_mcp.storage.graph import GraphBackend, _GRAPH_BASE
from office_mcp.utils.errors import (
    AuthRequiredError,
    BackendUnavailableError,
    InvalidUriError,
    PathTraversalError,
    PreconditionFailedError,
    StorageNotFoundError,
)


def _user(token: str | None) -> UserContext:
    return UserContext(user_id="demo-user", on_behalf_of_token=token)


@pytest_asyncio.fixture
async def backend() -> AsyncIterator[GraphBackend]:
    instance = GraphBackend()
    try:
        yield instance
    finally:
        await instance.aclose()


# ── URI / token guards ─────────────────────────────────────────────────


@pytest.mark.asyncio
async def test_read_without_token_raises(backend: GraphBackend) -> None:
    with pytest.raises(AuthRequiredError):
        await backend.read("od://x.txt", user_ctx=_user(None))


@pytest.mark.asyncio
async def test_path_traversal_rejected(backend: GraphBackend) -> None:
    with pytest.raises(PathTraversalError):
        await backend.read("od://../escape.txt", user_ctx=_user("AT"))


@pytest.mark.asyncio
async def test_invalid_uri_scheme_rejected(backend: GraphBackend) -> None:
    with pytest.raises(InvalidUriError):
        await backend.read("local://x.txt", user_ctx=_user("AT"))


@pytest.mark.asyncio
async def test_empty_uri_rejected(backend: GraphBackend) -> None:
    with pytest.raises(InvalidUriError):
        await backend.read("od://", user_ctx=_user("AT"))


# ── read ───────────────────────────────────────────────────────────────


@pytest.mark.asyncio
async def test_read_returns_bytes(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/Documents/q3.docx:/content",
        method="GET",
        content=b"hello world",
    )

    data = await backend.read(
        "od://Documents/q3.docx", user_ctx=_user("AT")
    )

    assert data == b"hello world"
    request = httpx_mock.get_requests()[-1]
    assert request.headers["Authorization"] == "Bearer AT"


@pytest.mark.asyncio
async def test_read_404_raises_not_found(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/missing.txt:/content",
        method="GET",
        status_code=404,
    )
    with pytest.raises(StorageNotFoundError):
        await backend.read("od://missing.txt", user_ctx=_user("AT"))


@pytest.mark.asyncio
async def test_read_401_raises_auth_required(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/x.txt:/content",
        method="GET",
        status_code=401,
    )
    with pytest.raises(AuthRequiredError):
        await backend.read("od://x.txt", user_ctx=_user("expired"))


# ── write ──────────────────────────────────────────────────────────────


@pytest.mark.asyncio
async def test_write_returns_metadata_with_etag(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/q.docx:/content",
        method="PUT",
        json={
            "size": 11,
            "lastModifiedDateTime": "2026-04-29T10:00:00Z",
            "eTag": '"{abc-123},1"',
            "file": {"mimeType": "application/octet-stream"},
        },
    )

    meta = await backend.write(
        "od://q.docx", b"hello world", user_ctx=_user("AT")
    )

    assert meta.size_bytes == 11
    assert meta.etag == '"{abc-123},1"'
    request = httpx_mock.get_requests()[-1]
    assert request.headers["Authorization"] == "Bearer AT"
    assert request.headers["Content-Type"] == "application/octet-stream"
    assert request.read() == b"hello world"


@pytest.mark.asyncio
async def test_write_with_if_match_attaches_header(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/q.docx:/content",
        method="PUT",
        json={
            "size": 1,
            "lastModifiedDateTime": "2026-04-29T10:00:00Z",
            "eTag": '"new-etag"',
            "file": {"mimeType": "text/plain"},
        },
    )

    await backend.write(
        "od://q.docx",
        b"x",
        user_ctx=_user("AT"),
        if_match='"old-etag"',
    )

    request = httpx_mock.get_requests()[-1]
    assert request.headers["If-Match"] == '"old-etag"'


@pytest.mark.asyncio
async def test_write_412_raises_precondition_failed(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/q.docx:/content",
        method="PUT",
        status_code=412,
    )
    with pytest.raises(PreconditionFailedError):
        await backend.write(
            "od://q.docx",
            b"x",
            user_ctx=_user("AT"),
            if_match='"stale"',
        )


@pytest.mark.asyncio
async def test_write_rejects_files_above_simple_put_cap(
    backend: GraphBackend,
) -> None:
    big = b"x" * (4 * 1024 * 1024 + 1)
    with pytest.raises(BackendUnavailableError, match="cap"):
        await backend.write("od://big.bin", big, user_ctx=_user("AT"))


@pytest.mark.asyncio
async def test_write_encodes_path_segments(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    """File names with spaces or accents must reach Graph URL-encoded."""
    expected_url = (
        f"{_GRAPH_BASE}/me/drive/root:/My%20Docs/r%C3%A9sum%C3%A9.docx:/content"
    )
    httpx_mock.add_response(
        url=expected_url,
        method="PUT",
        json={
            "size": 1,
            "lastModifiedDateTime": "2026-04-29T10:00:00Z",
            "eTag": '"e"',
            "file": {"mimeType": "application/octet-stream"},
        },
    )
    await backend.write(
        "od://My Docs/résumé.docx", b"x", user_ctx=_user("AT")
    )
    # If the URL didn't match the registered mock, pytest-httpx would
    # have raised; reaching here is the assertion.


# ── stat ───────────────────────────────────────────────────────────────


@pytest.mark.asyncio
async def test_stat_returns_metadata(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/q.docx",
        method="GET",
        json={
            "size": 42,
            "lastModifiedDateTime": "2026-04-29T11:00:00Z",
            "eTag": '"abc"',
            "file": {"mimeType": "application/vnd.x"},
        },
    )

    meta = await backend.stat("od://q.docx", user_ctx=_user("AT"))

    assert meta.uri == "od://q.docx"
    assert meta.size_bytes == 42
    assert meta.etag == '"abc"'
    assert meta.mime_type == "application/vnd.x"


@pytest.mark.asyncio
async def test_stat_missing_raises_not_found(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/missing.txt",
        method="GET",
        status_code=404,
    )
    with pytest.raises(StorageNotFoundError):
        await backend.stat("od://missing.txt", user_ctx=_user("AT"))


# ── delete ─────────────────────────────────────────────────────────────


@pytest.mark.asyncio
async def test_delete_returns_none_on_204(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/q.docx",
        method="DELETE",
        status_code=204,
    )
    await backend.delete("od://q.docx", user_ctx=_user("AT"))


@pytest.mark.asyncio
async def test_delete_404_raises_not_found(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/x.txt",
        method="DELETE",
        status_code=404,
    )
    with pytest.raises(StorageNotFoundError):
        await backend.delete("od://x.txt", user_ctx=_user("AT"))


# ── list (recursive BFS) ───────────────────────────────────────────────


@pytest.mark.asyncio
async def test_list_recursively_descends_into_folders(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    # Root: one file + one folder.
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root/children",
        method="GET",
        json={
            "value": [
                {
                    "name": "top.txt",
                    "size": 3,
                    "lastModifiedDateTime": "2026-04-29T10:00:00Z",
                    "eTag": '"e1"',
                    "file": {"mimeType": "text/plain"},
                },
                {
                    "name": "Sub",
                    "folder": {"childCount": 1},
                },
            ]
        },
    )
    # Sub folder: one nested file.
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/Sub:/children",
        method="GET",
        json={
            "value": [
                {
                    "name": "nested.txt",
                    "size": 4,
                    "lastModifiedDateTime": "2026-04-29T10:00:00Z",
                    "eTag": '"e2"',
                    "file": {"mimeType": "text/plain"},
                }
            ]
        },
    )

    results = [
        m.uri
        async for m in backend.list("od://", user_ctx=_user("AT"))
    ]

    assert results == ["od://top.txt", "od://Sub/nested.txt"]


@pytest.mark.asyncio
async def test_list_follows_odata_next_link(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    next_url = f"{_GRAPH_BASE}/me/drive/root/children?$skiptoken=42"
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root/children",
        method="GET",
        json={
            "value": [
                {
                    "name": "a.txt",
                    "size": 1,
                    "lastModifiedDateTime": "2026-04-29T10:00:00Z",
                    "eTag": '"e1"',
                    "file": {"mimeType": "text/plain"},
                }
            ],
            "@odata.nextLink": next_url,
        },
    )
    httpx_mock.add_response(
        url=next_url,
        method="GET",
        json={
            "value": [
                {
                    "name": "b.txt",
                    "size": 1,
                    "lastModifiedDateTime": "2026-04-29T10:00:00Z",
                    "eTag": '"e2"',
                    "file": {"mimeType": "text/plain"},
                }
            ]
        },
    )

    results = [
        m.uri
        async for m in backend.list("od://", user_ctx=_user("AT"))
    ]

    assert results == ["od://a.txt", "od://b.txt"]


# ── user scoping is implicit via the access token ──────────────────────


@pytest.mark.asyncio
async def test_each_user_token_maps_to_a_distinct_drive(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    """Two different tokens hit the same /me/drive endpoint but the
    server-side identity (and therefore the drive) is determined by
    the bearer alone — verified here by asserting the Authorization
    header attached to each request."""
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/x.txt:/content",
        method="GET",
        content=b"alice's bytes",
    )
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/x.txt:/content",
        method="GET",
        content=b"bob's bytes",
    )

    alice = await backend.read("od://x.txt", user_ctx=_user("AT-alice"))
    bob = await backend.read("od://x.txt", user_ctx=_user("AT-bob"))

    assert alice == b"alice's bytes"
    assert bob == b"bob's bytes"
    headers = [r.headers["Authorization"] for r in httpx_mock.get_requests()]
    assert headers == ["Bearer AT-alice", "Bearer AT-bob"]


# ── health check ───────────────────────────────────────────────────────


@pytest.mark.asyncio
async def test_health_check_ok_when_metadata_responds(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/$metadata",
        method="GET",
        status_code=200,
        text="<edmx />",
    )
    assert await backend.health_check() == "ok"


@pytest.mark.asyncio
async def test_health_check_raises_on_5xx(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/$metadata",
        method="GET",
        status_code=503,
    )
    with pytest.raises(RuntimeError):
        await backend.health_check()


# ── 5xx → BackendUnavailableError mapping ──────────────────────────────


@pytest.mark.asyncio
async def test_5xx_maps_to_backend_unavailable(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/x.txt:/content",
        method="GET",
        status_code=502,
        text="bad gateway",
    )
    with pytest.raises(BackendUnavailableError, match="502"):
        await backend.read("od://x.txt", user_ctx=_user("AT"))


# ── bootstrap registers `od` only when oauth is configured ─────────────


def test_bootstrap_registers_od_only_when_oauth_configured(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """The deploy registers GraphBackend conditionally — verify here so a
    refactor of build_app() can't silently drop the gate."""
    from unittest.mock import MagicMock

    from office_mcp.config import Settings
    from office_mcp.runtime import get_registry
    from office_mcp.server import build_app

    monkeypatch.setattr(
        "office_mcp.auth.oauth.msal.ConfidentialClientApplication",
        MagicMock(),
    )

    no_oauth = Settings(
        host="127.0.0.1",
        port=8081,
        log_level="warning",
        local_root=tmp_path,
    )
    build_app(no_oauth)
    schemes = dict(get_registry().items())
    assert "local" in schemes
    assert "od" not in schemes

    with_oauth = Settings(
        host="127.0.0.1",
        port=8082,
        log_level="warning",
        local_root=tmp_path,
        public_base_url="http://127.0.0.1:8082",
        azure_tenant_id="t",
        azure_client_id="c",
        azure_client_secret="s",  # noqa: S106 — dummy literal for tests
    )
    build_app(with_oauth)
    assert "od" in dict(get_registry().items())


