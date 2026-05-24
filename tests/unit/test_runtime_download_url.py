"""Unit tests for the download-URL selection helper.

The helper used to prefer a backend-native ``web_url`` (e.g. OneDrive's
``*.sharepoint.com`` URL) so the LLM could surface a one-click "open in
Word for the web" link. That turned out to be a trap: clicking the URL
starts a WOPI coauthoring session that locks the file against every
subsequent server-side write.

The current design asks Microsoft Graph for a read-only embed URL via
``POST /me/drive/items/<path>:/preview`` for ``od://`` URIs. Read-only
previews don't acquire the lock, so the user gets an in-browser viewer
*and* the chat-driven edits keep working. Local URIs (and Graph failures)
fall back to a signed ``/files/{token}`` URL.
"""

from __future__ import annotations

from collections.abc import Iterator

import pytest

from office_mcp.config import Settings
from office_mcp.context import UserContext
from office_mcp.runtime import download_url_for, set_registry, set_settings
from office_mcp.storage.base import FileMetadata


class _FakeBackend:
    """Stand-in for a backend whose ``preview_url`` returns a canned URL.

    Bypasses pytest-httpx so this file can stay focused on the
    dispatch + fallback logic; full Graph-level mocking lives in
    ``test_storage_graph.py``.
    """

    def __init__(self, *, preview: str | None | Exception) -> None:
        self._preview = preview

    async def preview_url(self, uri: str, *, user_ctx: UserContext) -> str | None:
        if isinstance(self._preview, Exception):
            raise self._preview
        return self._preview


class _FakeRegistry:
    """Minimal stand-in for :class:`StorageRegistry` — only ``get`` is exercised."""

    def __init__(self, backend: object) -> None:
        self._backend = backend

    def get(self, uri: str) -> object:
        return self._backend


@pytest.fixture(autouse=True)
def _restore_runtime_state() -> Iterator[None]:
    yield
    set_settings(None)
    set_registry(None)


def _meta(uri: str, *, web_url: str | None = None) -> FileMetadata:
    return FileMetadata(
        uri=uri,
        size_bytes=1,
        modified_at="2026-04-29T10:00:00Z",
        etag='"e"',
        mime_type=None,
        web_url=web_url,
    )


def _user() -> UserContext:
    return UserContext(
        user_id="alice",
        email="alice@example.com",
        on_behalf_of_token="user-AT",
    )


@pytest.mark.asyncio
async def test_returns_signed_url_for_local_uri() -> None:
    """Local backend never has a preview API — straight to the signed URL."""
    set_settings(
        Settings(
            signing_secret="s",  # noqa: S106 — dummy literal for tests
            download_base_url="https://office-mcp.test",
        )
    )
    set_registry(_FakeRegistry(object()))  # backend without preview_url

    url = await download_url_for(_meta("local://x.docx"), _user())

    assert url is not None
    assert url.startswith("https://office-mcp.test/files/")


@pytest.mark.asyncio
async def test_returns_preview_url_for_od_uri() -> None:
    """For ``od://`` URIs we hand the user the read-only embed link so
    the viewer doesn't acquire a WOPI write lock."""
    set_settings(
        Settings(
            signing_secret="s",  # noqa: S106 — dummy literal for tests
            download_base_url="https://office-mcp.test",
        )
    )
    set_registry(
        _FakeRegistry(_FakeBackend(preview="https://example-my.sharepoint.com/preview/abc"))
    )

    url = await download_url_for(_meta("od://q.docx"), _user())

    assert url == "https://example-my.sharepoint.com/preview/abc"


@pytest.mark.asyncio
async def test_falls_back_to_signed_url_when_preview_returns_none() -> None:
    """No token / Graph 4xx / malformed payload all collapse to ``None``
    inside ``preview_url`` — we must not propagate that to the caller,
    fall back to the signed-download path so the user still gets a link."""
    set_settings(
        Settings(
            signing_secret="s",  # noqa: S106 — dummy literal for tests
            download_base_url="https://office-mcp.test",
        )
    )
    set_registry(_FakeRegistry(_FakeBackend(preview=None)))

    url = await download_url_for(_meta("od://q.docx"), _user())

    assert url is not None
    assert url.startswith("https://office-mcp.test/files/")


@pytest.mark.asyncio
async def test_falls_back_to_signed_url_when_preview_raises() -> None:
    """A network error or other exception inside the preview call must
    not propagate — same fallback as a clean ``None``."""
    set_settings(
        Settings(
            signing_secret="s",  # noqa: S106 — dummy literal for tests
            download_base_url="https://office-mcp.test",
        )
    )
    set_registry(_FakeRegistry(_FakeBackend(preview=RuntimeError("boom"))))

    url = await download_url_for(_meta("od://q.docx"), _user())

    assert url is not None
    assert url.startswith("https://office-mcp.test/files/")


@pytest.mark.asyncio
async def test_returns_none_when_signing_disabled_and_no_preview() -> None:
    """Without a signing secret + download_base_url AND no preview URL,
    we have no link to offer — returning ``None`` is honest. Tools omit
    the ``download_url`` field rather than send a broken link."""
    set_settings(Settings())  # no signing config
    set_registry(_FakeRegistry(_FakeBackend(preview=None)))

    assert await download_url_for(_meta("local://x.docx"), _user()) is None
    assert await download_url_for(_meta("od://q.docx"), _user()) is None
