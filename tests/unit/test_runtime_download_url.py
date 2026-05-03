"""Unit tests for the download-URL selection helper.

The helper picks between a backend-native URL (``meta.web_url`` — for
example OneDrive's ``webUrl``) and a locally signed ``/files/{token}``
fallback, so the LLM always surfaces the most useful link to the user.
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


def test_prefers_backend_web_url_when_present() -> None:
    set_settings(
        Settings(
            signing_secret="s",  # noqa: S106 — dummy literal for tests
            download_base_url="https://office-mcp.test",
        )
    )
    web_url = "https://example-my.sharepoint.com/personal/q.docx"

    assert (
        download_url_for(_meta("od://q.docx", web_url=web_url), "yacine")
        == web_url
    )


def test_falls_back_to_signed_url_when_no_web_url() -> None:
    set_settings(
        Settings(
            signing_secret="s",  # noqa: S106 — dummy literal for tests
            download_base_url="https://office-mcp.test",
        )
    )

    url = download_url_for(_meta("local://x.docx"), "yacine")

    assert url is not None
    assert url.startswith("https://office-mcp.test/files/")


def test_returns_none_when_signing_disabled_and_no_web_url() -> None:
    set_settings(Settings())  # no signing_secret + no download_base_url

    assert download_url_for(_meta("local://x.docx"), "yacine") is None


def test_returns_web_url_even_when_signing_disabled() -> None:
    """The web URL doesn't need our signing infra to be useful."""
    set_settings(Settings())
    web_url = "https://example-my.sharepoint.com/personal/q.docx"

    assert (
        download_url_for(_meta("od://q.docx", web_url=web_url), "yacine")
        == web_url
    )
