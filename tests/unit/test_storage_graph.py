"""Unit tests for the Microsoft Graph (OneDrive) storage backend.

We use ``pytest-httpx`` to stub Graph responses so the suite never
contacts ``graph.microsoft.com``. The tests focus on the URL shape we
emit, the headers we attach, and the way we map provider statuses to
office-mcp errors.
"""

from __future__ import annotations

from collections.abc import AsyncIterator
from datetime import UTC, datetime, timedelta
from pathlib import Path

import pytest
import pytest_asyncio
from pytest_httpx import HTTPXMock

from office_mcp.context import UserContext
from office_mcp.storage.graph import _GRAPH_BASE, GraphBackend
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


def _iso_age(seconds_ago: float) -> str:
    """ISO-8601 ``Z`` timestamp this many seconds in the past."""
    ts = datetime.now(UTC) - timedelta(seconds=seconds_ago)
    return ts.isoformat().replace("+00:00", "Z")


#: Fixed pre-signed upload URL used by ``test_write_*`` mocks.
#: Production gets a fresh per-request URL from Graph; for tests a
#: stable value lets pytest-httpx match the second-leg PUT exactly.
_TEST_UPLOAD_URL = "https://upload.example/token-abc"


def _mock_upload_session_ok(
    httpx_mock: HTTPXMock,
    path: str,
    *,
    final_json: dict[str, object] | None = None,
    final_status: int = 201,
) -> None:
    """Register the happy-path two-step upload-session flow.

    Tests that only need a successful write can call this once and
    forget about the protocol shape.
    """
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/{path}:/createUploadSession",
        method="POST",
        json={
            "uploadUrl": _TEST_UPLOAD_URL,
            "expirationDateTime": "2099-01-01T00:00:00Z",
        },
    )
    httpx_mock.add_response(
        url=_TEST_UPLOAD_URL,
        method="PUT",
        status_code=final_status,
        json=final_json
        or {
            "size": 1,
            "lastModifiedDateTime": "2026-04-29T10:00:00Z",
            "eTag": '"e"',
            "file": {"mimeType": "application/octet-stream"},
        },
    )


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
async def test_empty_path_segment_rejected(backend: GraphBackend) -> None:
    """``od://foo//bar`` would 404 confusingly at Graph; reject early."""
    with pytest.raises(PathTraversalError):
        await backend.read("od://foo//bar.txt", user_ctx=_user("AT"))


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
async def test_read_follows_graph_302_to_cdn(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    """Real Graph 302s the bytes onto a short-lived CDN URL.

    follow_redirects=True is the production default; this test locks
    that behavior so a refactor can't silently flip it to False.
    """
    cdn_url = "https://cdn.example.com/blob/abc"
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/q.docx:/content",
        method="GET",
        status_code=302,
        headers={"Location": cdn_url},
    )
    httpx_mock.add_response(url=cdn_url, method="GET", content=b"payload")

    data = await backend.read("od://q.docx", user_ctx=_user("AT"))

    assert data == b"payload"


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
    _mock_upload_session_ok(
        httpx_mock,
        "q.docx",
        final_json={
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

    create_req, upload_req = httpx_mock.get_requests()
    # Session-create call carries the bearer; the upload URL is
    # pre-signed so the chunk PUT must NOT include Authorization
    # (Microsoft Graph docs are explicit about this).
    assert create_req.method == "POST"
    assert create_req.headers["Authorization"] == "Bearer AT"
    assert upload_req.method == "PUT"
    assert "authorization" not in {k.lower() for k in upload_req.headers}
    assert upload_req.headers["Content-Length"] == "11"
    assert upload_req.headers["Content-Range"] == "bytes 0-10/11"
    assert upload_req.read() == b"hello world"


@pytest.mark.asyncio
async def test_write_prefers_etag_over_ctag(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    """Optimistic concurrency uses ``eTag`` (item-wide), not ``cTag``
    (content-only). Locks the contract so a refactor can't silently
    switch which header drives ``If-Match``."""
    _mock_upload_session_ok(
        httpx_mock,
        "q.docx",
        final_json={
            "size": 1,
            "lastModifiedDateTime": "2026-04-29T10:00:00Z",
            "eTag": '"correct"',
            "cTag": '"wrong"',
            "file": {"mimeType": "application/octet-stream"},
        },
    )

    meta = await backend.write("od://q.docx", b"x", user_ctx=_user("AT"))

    assert meta.etag == '"correct"'


@pytest.mark.asyncio
async def test_write_with_if_match_attaches_header(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    """``If-Match`` rides the final ranged PUT, not the
    ``createUploadSession`` POST — that's where Graph evaluates the
    precondition for the actual content swap."""
    _mock_upload_session_ok(httpx_mock, "q.docx")

    await backend.write(
        "od://q.docx",
        b"x",
        user_ctx=_user("AT"),
        if_match='"old-etag"',
    )

    create_req, upload_req = httpx_mock.get_requests()
    assert "If-Match" not in create_req.headers
    assert upload_req.headers["If-Match"] == '"old-etag"'


@pytest.mark.asyncio
async def test_write_412_raises_precondition_failed(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    """412 from the upload-URL PUT (stale If-Match) must surface as
    ``PreconditionFailedError`` — the optimistic-concurrency contract
    stays identical across the simple-PUT and upload-session paths.
    """
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/q.docx:/createUploadSession",
        method="POST",
        json={"uploadUrl": _TEST_UPLOAD_URL},
    )
    httpx_mock.add_response(
        url=_TEST_UPLOAD_URL,
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
async def test_write_retries_through_brief_resource_lock(
    backend: GraphBackend, httpx_mock: HTTPXMock, monkeypatch: pytest.MonkeyPatch
) -> None:
    """OneDrive can briefly 423 the ``createUploadSession`` POST too,
    when SharePoint is still propagating a recent change. The retry
    helper absorbs that the same way it did for simple-PUT, so the LLM
    never sees a transient lock as a hard failure (it would otherwise
    clone the file under a v2 name — observed in production)."""
    # Speed up the test by collapsing the backoff schedule.
    monkeypatch.setattr(
        "office_mcp.storage.graph._RETRY_LOCKED_DELAYS_S", (0.0, 0.0)
    )
    create_url = f"{_GRAPH_BASE}/me/drive/root:/q.docx:/createUploadSession"
    httpx_mock.add_response(url=create_url, method="POST", status_code=423)
    # After the first 423 the backend probes lastModifiedDateTime to
    # distinguish the brief post-write lock (this case) from an external
    # coauthoring lock — fresh mtime means "transient, keep retrying".
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/q.docx",
        method="GET",
        json={
            "size": 1,
            "lastModifiedDateTime": _iso_age(5),
            "eTag": '"old"',
            "file": {"mimeType": "application/octet-stream"},
        },
    )
    httpx_mock.add_response(url=create_url, method="POST", status_code=423)
    httpx_mock.add_response(
        url=create_url,
        method="POST",
        json={"uploadUrl": _TEST_UPLOAD_URL},
    )
    httpx_mock.add_response(
        url=_TEST_UPLOAD_URL,
        method="PUT",
        json={
            "size": 1,
            "lastModifiedDateTime": "2026-04-29T10:00:00Z",
            "eTag": '"new"',
            "file": {"mimeType": "application/octet-stream"},
        },
    )

    meta = await backend.write("od://q.docx", b"x", user_ctx=_user("AT"))

    assert meta.etag == '"new"'
    # 3 POSTs + 1 diagnostic GET + 1 final PUT = 5 HTTP calls.
    assert len(httpx_mock.get_requests()) == 5


@pytest.mark.asyncio
async def test_write_surfaces_unavailable_when_lock_persists(
    backend: GraphBackend, httpx_mock: HTTPXMock, monkeypatch: pytest.MonkeyPatch
) -> None:
    """If retries don't clear the lock we surface the error instead of
    looping forever."""
    monkeypatch.setattr(
        "office_mcp.storage.graph._RETRY_LOCKED_DELAYS_S", (0.0, 0.0)
    )
    create_url = f"{_GRAPH_BASE}/me/drive/root:/q.docx:/createUploadSession"
    for _ in range(3):  # initial + 2 retries
        httpx_mock.add_response(url=create_url, method="POST", status_code=423)
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/q.docx",
        method="GET",
        json={
            "size": 1,
            "lastModifiedDateTime": _iso_age(5),
            "eTag": '"old"',
            "file": {"mimeType": "application/octet-stream"},
        },
    )

    with pytest.raises(BackendUnavailableError, match="still indexing"):
        await backend.write("od://q.docx", b"x", user_ctx=_user("AT"))


@pytest.mark.asyncio
async def test_write_short_circuits_on_external_coauthoring_lock(
    backend: GraphBackend, httpx_mock: HTTPXMock, monkeypatch: pytest.MonkeyPatch
) -> None:
    """A 423 on a file modified several minutes ago can only come from
    an external coauthoring session (Word Online, OneDrive desktop) —
    no server-side retry will release it. We bail out fast with an
    actionable message instead of burning the retry budget for nothing.

    Upload-session is supposed to coexist with WOPI sessions; in
    practice we still keep this safety net for the rare cases where
    Graph rejects the session create itself.
    """
    monkeypatch.setattr(
        "office_mcp.storage.graph._RETRY_LOCKED_DELAYS_S", (0.0, 0.0)
    )
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/q.docx:/createUploadSession",
        method="POST",
        status_code=423,
    )
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/q.docx",
        method="GET",
        json={
            "size": 1,
            "lastModifiedDateTime": _iso_age(600),
            "eTag": '"old"',
            "file": {"mimeType": "application/octet-stream"},
        },
    )

    with pytest.raises(BackendUnavailableError, match="Word Online"):
        await backend.write("od://q.docx", b"x", user_ctx=_user("AT"))

    assert len(httpx_mock.get_requests()) == 2


@pytest.mark.asyncio
async def test_write_falls_back_to_retry_when_lock_probe_fails(
    backend: GraphBackend, httpx_mock: HTTPXMock, monkeypatch: pytest.MonkeyPatch
) -> None:
    """A failing classifier probe (404, network blip, malformed JSON)
    must not break the normal retry path — we just fall back to the
    transient assumption."""
    monkeypatch.setattr(
        "office_mcp.storage.graph._RETRY_LOCKED_DELAYS_S", (0.0, 0.0)
    )
    create_url = f"{_GRAPH_BASE}/me/drive/root:/q.docx:/createUploadSession"
    httpx_mock.add_response(url=create_url, method="POST", status_code=423)
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/q.docx",
        method="GET",
        status_code=404,
    )
    httpx_mock.add_response(
        url=create_url,
        method="POST",
        json={"uploadUrl": _TEST_UPLOAD_URL},
    )
    httpx_mock.add_response(
        url=_TEST_UPLOAD_URL,
        method="PUT",
        json={
            "size": 1,
            "lastModifiedDateTime": "2026-04-29T10:00:00Z",
            "eTag": '"new"',
            "file": {"mimeType": "application/octet-stream"},
        },
    )

    meta = await backend.write("od://q.docx", b"x", user_ctx=_user("AT"))
    assert meta.etag == '"new"'


@pytest.mark.asyncio
async def test_write_rejects_zero_byte_uploads(backend: GraphBackend) -> None:
    """Graph rejects ``Content-Range: bytes 0--1/0`` and so do many
    intermediaries; we never expect a real zero-byte write from the
    engines, so fail closed at the API boundary."""
    with pytest.raises(BackendUnavailableError, match="zero bytes"):
        await backend.write("od://empty.docx", b"", user_ctx=_user("AT"))


@pytest.mark.asyncio
async def test_write_rejects_files_above_upload_session_cap(
    backend: GraphBackend,
) -> None:
    """Single-shot upload-session caps at 60 MiB. Beyond that we'd need
    chunked uploads — out of scope until a real use case lands."""
    big = b"x" * (60 * 1024 * 1024 + 1)
    with pytest.raises(BackendUnavailableError, match="cap"):
        await backend.write("od://big.bin", big, user_ctx=_user("AT"))


@pytest.mark.asyncio
async def test_write_encodes_path_segments(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    """File names with spaces or accents must reach Graph URL-encoded
    on the session-create POST (the upload URL is opaque after that)."""
    expected_url = (
        f"{_GRAPH_BASE}/me/drive/root:"
        f"/My%20Docs/r%C3%A9sum%C3%A9.docx:/createUploadSession"
    )
    httpx_mock.add_response(
        url=expected_url,
        method="POST",
        json={"uploadUrl": _TEST_UPLOAD_URL},
    )
    httpx_mock.add_response(
        url=_TEST_UPLOAD_URL,
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


# ── default-folder normalization (writes only) ────────────────────────


@pytest.mark.asyncio
async def test_write_rewrites_bare_uri_into_default_folder(
    httpx_mock: HTTPXMock,
) -> None:
    """``od://foo.docx`` (no folder) lands under the configured default
    folder rather than at the OneDrive root, so LLM-generated files end
    up in a predictable place."""
    backend = GraphBackend(default_folder="AI Generated")
    try:
        _mock_upload_session_ok(
            httpx_mock,
            "AI%20Generated/foo.docx",
            final_json={
                "size": 1,
                "lastModifiedDateTime": "2026-04-29T10:00:00Z",
                "eTag": '"e"',
                "file": {"mimeType": "application/octet-stream"},
                "webUrl": "https://example-my.sharepoint.com/personal/foo.docx",
            },
        )

        meta = await backend.write(
            "od://foo.docx", b"x", user_ctx=_user("AT")
        )

        assert meta.uri == "od://AI Generated/foo.docx"
        assert meta.web_url == "https://example-my.sharepoint.com/personal/foo.docx"
    finally:
        await backend.aclose()


@pytest.mark.asyncio
async def test_write_passes_explicit_folder_through_unchanged(
    httpx_mock: HTTPXMock,
) -> None:
    """A URI with an explicit folder must not be rewritten — only bare
    names trigger normalization."""
    backend = GraphBackend(default_folder="AI Generated")
    try:
        _mock_upload_session_ok(httpx_mock, "Documents/q3.docx")

        meta = await backend.write(
            "od://Documents/q3.docx", b"x", user_ctx=_user("AT")
        )

        assert meta.uri == "od://Documents/q3.docx"
    finally:
        await backend.aclose()


@pytest.mark.asyncio
async def test_write_default_folder_disabled_keeps_root_writes(
    httpx_mock: HTTPXMock,
) -> None:
    """Operators can opt out by passing an empty default_folder."""
    backend = GraphBackend(default_folder="")
    try:
        _mock_upload_session_ok(httpx_mock, "foo.docx")

        meta = await backend.write(
            "od://foo.docx", b"x", user_ctx=_user("AT")
        )

        assert meta.uri == "od://foo.docx"
    finally:
        await backend.aclose()


@pytest.mark.asyncio
async def test_write_captures_web_url_when_present(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    _mock_upload_session_ok(
        httpx_mock,
        "q.docx",
        final_json={
            "size": 1,
            "lastModifiedDateTime": "2026-04-29T10:00:00Z",
            "eTag": '"e"',
            "file": {"mimeType": "application/octet-stream"},
            "webUrl": "https://example-my.sharepoint.com/personal/q.docx",
        },
    )

    meta = await backend.write("od://q.docx", b"x", user_ctx=_user("AT"))

    assert meta.web_url == "https://example-my.sharepoint.com/personal/q.docx"


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


@pytest.mark.asyncio
async def test_delete_retries_through_brief_resource_lock(
    backend: GraphBackend, httpx_mock: HTTPXMock, monkeypatch: pytest.MonkeyPatch
) -> None:
    """DELETE hits the same post-write 423 lock as PUT — the LLM
    asking to remove a file it just created would otherwise see a hard
    BackendUnavailableError. Reuses ``_retry_transient`` like ``write``."""
    monkeypatch.setattr(
        "office_mcp.storage.graph._RETRY_LOCKED_DELAYS_S", (0.0, 0.0)
    )
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/q.docx",
        method="DELETE",
        status_code=423,
    )
    # Diagnostic probe after the first 423 — fresh mtime keeps the
    # classification "transient" so the retry loop continues.
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/q.docx",
        method="GET",
        json={
            "size": 1,
            "lastModifiedDateTime": _iso_age(5),
            "eTag": '"old"',
            "file": {"mimeType": "application/octet-stream"},
        },
    )
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/q.docx",
        method="DELETE",
        status_code=423,
    )
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/q.docx",
        method="DELETE",
        status_code=204,
    )

    await backend.delete("od://q.docx", user_ctx=_user("AT"))
    # 3 DELETEs + 1 diagnostic GET = 4 HTTP calls.
    assert len(httpx_mock.get_requests()) == 4


@pytest.mark.asyncio
async def test_delete_surfaces_unavailable_when_lock_persists(
    backend: GraphBackend, httpx_mock: HTTPXMock, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(
        "office_mcp.storage.graph._RETRY_LOCKED_DELAYS_S", (0.0, 0.0)
    )
    for _ in range(3):  # initial + 2 retries
        httpx_mock.add_response(
            url=f"{_GRAPH_BASE}/me/drive/root:/q.docx",
            method="DELETE",
            status_code=423,
        )
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/q.docx",
        method="GET",
        json={
            "size": 1,
            "lastModifiedDateTime": _iso_age(5),  # fresh → transient
            "eTag": '"old"',
            "file": {"mimeType": "application/octet-stream"},
        },
    )

    with pytest.raises(BackendUnavailableError, match="still indexing"):
        await backend.delete("od://q.docx", user_ctx=_user("AT"))


@pytest.mark.asyncio
async def test_delete_short_circuits_on_external_coauthoring_lock(
    backend: GraphBackend, httpx_mock: HTTPXMock, monkeypatch: pytest.MonkeyPatch
) -> None:
    """The user's open Word Online tab on the file we're trying to
    delete holds the lock. Server-side retry can't help — fail fast
    with the actionable message so the LLM tells the user what to do.
    """
    monkeypatch.setattr(
        "office_mcp.storage.graph._RETRY_LOCKED_DELAYS_S", (0.0, 0.0)
    )
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/q.docx",
        method="DELETE",
        status_code=423,
    )
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/q.docx",
        method="GET",
        json={
            "size": 1,
            "lastModifiedDateTime": _iso_age(600),  # 10 min → external
            "eTag": '"old"',
            "file": {"mimeType": "application/octet-stream"},
        },
    )

    with pytest.raises(BackendUnavailableError, match="Word Online"):
        await backend.delete("od://q.docx", user_ctx=_user("AT"))

    assert len(httpx_mock.get_requests()) == 2


# ── preview (read-only viewer URL) ─────────────────────────────────────


@pytest.mark.asyncio
async def test_preview_url_returns_get_url_from_graph(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    """The read-only viewer URL replaces meta.web_url in user-facing
    responses — clicking it opens Office for the web in view mode and
    crucially does NOT acquire the WOPI write lock."""
    embed = (
        "https://example-my.sharepoint.com/personal/_layouts/15/"
        "Doc.aspx?sourcedoc=ABC&action=embedview"
    )
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/q.docx:/preview",
        method="POST",
        json={"getUrl": embed, "postUrl": None},
    )

    url = await backend.preview_url("od://q.docx", user_ctx=_user("AT"))

    assert url == embed


@pytest.mark.asyncio
async def test_preview_url_returns_none_without_token(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    """No anonymous /preview — without a user token we can't call it,
    so report None and let the caller fall back to a signed URL."""
    url = await backend.preview_url("od://q.docx", user_ctx=_user(None))

    assert url is None
    # No HTTP call should have been issued.
    assert httpx_mock.get_requests() == []


@pytest.mark.asyncio
async def test_preview_url_returns_none_on_graph_404(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    """File doesn't exist (or is otherwise unreachable) — return None
    rather than raising; the caller falls back to a signed URL."""
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/missing.docx:/preview",
        method="POST",
        status_code=404,
    )

    url = await backend.preview_url("od://missing.docx", user_ctx=_user("AT"))

    assert url is None


@pytest.mark.asyncio
async def test_preview_url_returns_none_on_malformed_response(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    """Graph said 200 but the body has no usable ``getUrl`` — treat as
    failure so the caller can degrade gracefully."""
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/q.docx:/preview",
        method="POST",
        json={"unexpected": "shape"},
    )

    url = await backend.preview_url("od://q.docx", user_ctx=_user("AT"))

    assert url is None


@pytest.mark.asyncio
async def test_preview_url_rejects_path_traversal(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    """URI parse errors short-circuit BEFORE hitting Graph — the URI is
    bad, no preview makes sense."""
    url = await backend.preview_url("od://../escape.docx", user_ctx=_user("AT"))

    assert url is None
    assert httpx_mock.get_requests() == []


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
async def test_each_request_forwards_its_bearer_token(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    """The actual cross-user isolation is a Graph-side property of the
    access token; pytest-httpx can only see what we send. So this test
    asserts the narrower invariant we control: each call attaches the
    user's own bearer token verbatim, never reusing or mixing them."""
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
    backend: GraphBackend,
    httpx_mock: HTTPXMock,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A 5xx that persists through the retry budget must still surface
    as ``BackendUnavailableError`` so the LLM doesn't silently swallow
    it. We collapse the backoff schedule to keep the test fast."""
    monkeypatch.setattr(
        "office_mcp.storage.graph._RETRY_5XX_DELAYS_S", (0.0, 0.0, 0.0)
    )
    url = f"{_GRAPH_BASE}/me/drive/root:/x.txt:/content"
    for _ in range(4):  # initial attempt + 3 retries
        httpx_mock.add_response(
            url=url, method="GET", status_code=502, text="bad gateway"
        )
    with pytest.raises(BackendUnavailableError, match="502"):
        await backend.read("od://x.txt", user_ctx=_user("AT"))


@pytest.mark.asyncio
async def test_429_above_cap_surfaces_retry_after(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    """Retry-After above ``_MAX_RETRY_AFTER_S`` is too long to absorb
    inline (would blow past LibreChat's 120 s tool timeout). We surface
    the 429 with the hint so the LLM tells the user when to retry."""
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/x.txt:/content",
        method="GET",
        status_code=429,
        headers={"Retry-After": "60"},
    )
    with pytest.raises(BackendUnavailableError, match="retry after 60"):
        await backend.read("od://x.txt", user_ctx=_user("AT"))


@pytest.mark.asyncio
async def test_5xx_retried_then_succeeds(
    backend: GraphBackend,
    httpx_mock: HTTPXMock,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A handful of transient 5xx that clear within the retry budget
    must not surface to the LLM at all — that's the whole point of the
    retry. Mirrors Microsoft's documented best practice for upload
    sessions (retry 500/502/503/504 with exponential backoff)."""
    monkeypatch.setattr(
        "office_mcp.storage.graph._RETRY_5XX_DELAYS_S", (0.0, 0.0, 0.0)
    )
    url = f"{_GRAPH_BASE}/me/drive/root:/x.txt:/content"
    httpx_mock.add_response(url=url, method="GET", status_code=503)
    httpx_mock.add_response(url=url, method="GET", status_code=502)
    httpx_mock.add_response(url=url, method="GET", content=b"finally")

    data = await backend.read("od://x.txt", user_ctx=_user("AT"))

    assert data == b"finally"


@pytest.mark.asyncio
async def test_504_retried_alongside_other_5xx(
    backend: GraphBackend,
    httpx_mock: HTTPXMock,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """504 Gateway Timeout is part of the retriable set per MS upload-
    session best practices. Lock the contract so a future refactor of
    ``_RETRIABLE_5XX`` doesn't drop it silently."""
    monkeypatch.setattr(
        "office_mcp.storage.graph._RETRY_5XX_DELAYS_S", (0.0, 0.0, 0.0)
    )
    url = f"{_GRAPH_BASE}/me/drive/root:/x.txt:/content"
    httpx_mock.add_response(url=url, method="GET", status_code=504)
    httpx_mock.add_response(url=url, method="GET", content=b"ok")

    data = await backend.read("od://x.txt", user_ctx=_user("AT"))
    assert data == b"ok"


@pytest.mark.asyncio
async def test_429_within_cap_is_honored_and_retried(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    """A small Retry-After value is honored inline: the backend sleeps
    that long, retries, and the LLM never sees the 429. Avoids the
    failure mode where the LLM gives up on a tool call that would have
    succeeded after a one-second wait."""
    url = f"{_GRAPH_BASE}/me/drive/root:/x.txt:/content"
    httpx_mock.add_response(
        url=url,
        method="GET",
        status_code=429,
        headers={"Retry-After": "0"},
    )
    httpx_mock.add_response(url=url, method="GET", content=b"served")

    data = await backend.read("od://x.txt", user_ctx=_user("AT"))

    assert data == b"served"
    assert len(httpx_mock.get_requests()) == 2


@pytest.mark.asyncio
async def test_429_without_retry_after_surfaces_immediately(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    """No Retry-After header means we can't safely guess a delay —
    surface the 429 with a generic message so the LLM backs off."""
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/x.txt:/content",
        method="GET",
        status_code=429,
    )
    with pytest.raises(BackendUnavailableError, match="throttled"):
        await backend.read("od://x.txt", user_ctx=_user("AT"))


@pytest.mark.asyncio
async def test_423_permanent_body_code_short_circuits(
    backend: GraphBackend,
    httpx_mock: HTTPXMock,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """When Graph returns a documented permanent lock code in the
    response body, we trust it over the mtime heuristic and bail out
    immediately. No diagnostic probe should fire — the body already
    told us this won't clear on retry."""
    monkeypatch.setattr(
        "office_mcp.storage.graph._RETRY_LOCKED_DELAYS_S", (0.0, 0.0)
    )
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/q.docx:/createUploadSession",
        method="POST",
        status_code=423,
        json={
            "error": {
                "code": "wopiLockMismatch",
                "message": "The file is locked by another WOPI client.",
            }
        },
    )

    with pytest.raises(BackendUnavailableError, match="wopiLockMismatch"):
        await backend.write("od://q.docx", b"x", user_ctx=_user("AT"))

    # No diagnostic GET should have been issued — the body code is
    # authoritative.
    assert len(httpx_mock.get_requests()) == 1


@pytest.mark.asyncio
async def test_423_transient_body_code_skips_mtime_probe(
    backend: GraphBackend,
    httpx_mock: HTTPXMock,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A documented transient code keeps us in the retry loop without
    paying for a mtime probe. The probe only fires when we have no
    body-level signal to lean on."""
    monkeypatch.setattr(
        "office_mcp.storage.graph._RETRY_LOCKED_DELAYS_S", (0.0, 0.0)
    )
    create_url = f"{_GRAPH_BASE}/me/drive/root:/q.docx:/createUploadSession"
    httpx_mock.add_response(
        url=create_url,
        method="POST",
        status_code=423,
        json={"error": {"code": "resourceLocked", "message": "transient"}},
    )
    httpx_mock.add_response(
        url=create_url,
        method="POST",
        json={"uploadUrl": _TEST_UPLOAD_URL},
    )
    httpx_mock.add_response(
        url=_TEST_UPLOAD_URL,
        method="PUT",
        json={
            "size": 1,
            "lastModifiedDateTime": "2026-04-29T10:00:00Z",
            "eTag": '"new"',
            "file": {"mimeType": "application/octet-stream"},
        },
    )

    meta = await backend.write("od://q.docx", b"x", user_ctx=_user("AT"))

    assert meta.etag == '"new"'
    # 2 POSTs + 1 PUT = 3 — no diagnostic GET (body code told us
    # everything we needed).
    assert len(httpx_mock.get_requests()) == 3


@pytest.mark.asyncio
async def test_507_surfaces_quota_message(
    backend: GraphBackend, httpx_mock: HTTPXMock
) -> None:
    """OneDrive quota exhaustion is permanent until freed — clearer
    than the generic "backend unavailable" string. 507 now surfaces
    on the upload-session POST since that's our first write contact."""
    httpx_mock.add_response(
        url=f"{_GRAPH_BASE}/me/drive/root:/big.bin:/createUploadSession",
        method="POST",
        status_code=507,
    )
    with pytest.raises(BackendUnavailableError, match="quota"):
        await backend.write("od://big.bin", b"x", user_ctx=_user("AT"))


# ── bootstrap registers `od` only when oauth is configured ─────────────


def test_bootstrap_registers_only_one_backend_at_a_time(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Production deploys register exactly one backend per build_app().

    When Azure AD is configured (``oauth_configured``) only ``od://`` is
    registered — local:// is intentionally hidden so the LLM never has
    to disambiguate where files should live. When Azure is absent
    (dev mode) we fall back to LocalBackend so the server still has a
    usable backend.

    Relies on ``_bootstrap`` always rebuilding the registry on each call
    (the ``_BOOTSTRAPPED`` early-return is *after* the registry block).
    If that invariant moves, this test starts passing for the wrong
    reason — keep them aligned.
    """
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
    schemes = dict(get_registry().items())
    assert "od" in schemes
    assert "local" not in schemes


