"""Unit tests for the download-URL selection helper.

The helper used to prefer a backend-native ``web_url`` (e.g. OneDrive's
``*.sharepoint.com`` URL) so the LLM could surface a one-click "open in
Word for the web" link. That turned out to be a trap: clicking the URL
starts a WOPI coauthoring session that locks the file against every
subsequent server-side write. The helper now *always* returns a signed
``/files/{token}`` URL so the click downloads the file instead of
opening it in Office for the web.
"""

from __future__ import annotations

import pytest

from office_mcp.config import Settings
from office_mcp.runtime import download_url_for, set_settings
from office_mcp.storage.base import FileMetadata


@pytest.fixture(autouse=True)
def _restore_settings() -> object:
    yield
    set_settings(None)


def _meta(uri: str, *, web_url: str | None = None) -> FileMetadata:
    return FileMetadata(
        uri=uri,
        size_bytes=1,
        modified_at="2026-04-29T10:00:00Z",
        etag='"e"',
        mime_type=None,
        web_url=web_url,
    )


def test_returns_signed_url_for_local_uri() -> None:
    set_settings(
        Settings(
            signing_secret="s",  # noqa: S106 — dummy literal for tests
            download_base_url="https://office-mcp.test",
        )
    )

    url = download_url_for(_meta("local://x.docx"), "yacine")

    assert url is not None
    assert url.startswith("https://office-mcp.test/files/")


def test_ignores_web_url_for_onedrive_uri() -> None:
    """The OneDrive ``webUrl`` is deliberately NOT returned: clicking it
    opens Word for the web, which starts a WOPI coauthoring session and
    locks the file against further writes. We always route through the
    signed download endpoint instead."""
    set_settings(
        Settings(
            signing_secret="s",  # noqa: S106 — dummy literal for tests
            download_base_url="https://office-mcp.test",
        )
    )
    web_url = "https://example-my.sharepoint.com/personal/q.docx"

    url = download_url_for(_meta("od://q.docx", web_url=web_url), "yacine")

    assert url is not None
    assert url != web_url
    assert url.startswith("https://office-mcp.test/files/")


def test_returns_none_when_signing_disabled() -> None:
    """Without a signing secret + download_base_url we have no URL to
    offer — even if the backend has a ``web_url``, surfacing it would
    re-introduce the coauthoring-lock trap."""
    set_settings(Settings())
    web_url = "https://example-my.sharepoint.com/personal/q.docx"

    assert download_url_for(_meta("local://x.docx"), "yacine") is None
    assert (
        download_url_for(_meta("od://q.docx", web_url=web_url), "yacine")
        is None
    )
