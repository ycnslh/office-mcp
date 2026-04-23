"""Process-wide runtime state: the FastMCP instance, the storage registry,
and the active :class:`Settings`.

Kept in its own module so tool/resource modules can import the ``mcp``
instance and shared state without creating a cycle through
``office_mcp.server``.
"""

from __future__ import annotations

from fastmcp import FastMCP

from office_mcp.config import Settings
from office_mcp.storage import StorageRegistry
from office_mcp.utils import signing

_INSTRUCTIONS = """\
office-mcp creates and edits Word, Excel and PowerPoint documents.

ADDRESSING FILES
  Every tool takes a `file_uri` like `local://path/to/file.ext`.
  In this deployment, ONLY the `local://` scheme is wired — do not invent
  `s3://`, `az://`, `sp://` or `od://` URIs (Phases 7-8 will add them).
  Path components are scoped per user, so a leading `/` is unnecessary.

WORKFLOW
  1. Before creating a document, consult the matching `office://rules/<format>`
     resource (docx, xlsx, pptx) for formatting conventions.
  2. Compose the document with the granular tools (create → write/append →
     read to verify).
  3. Every mutating tool returns a `download_url` — render it as a Markdown
     link in your answer so the user can click to download the file
     directly from the conversation. Example:
       `→ [q3.docx](http://host:18080/files/...)`
  4. Always echo the resulting `file_uri` alongside the download link.

ERROR HANDLING
  If a tool returns an UnknownSchemeError or InvalidUriError, RE-CALL the
  tool with `local://...` instead of guessing another scheme.
"""

mcp: FastMCP = FastMCP("office-mcp", instructions=_INSTRUCTIONS)

_registry: StorageRegistry | None = None


def get_registry() -> StorageRegistry:
    """Return the process-wide storage registry or raise if not initialised."""
    if _registry is None:
        raise RuntimeError(
            "Storage registry not initialised — build_app() or set_registry() "
            "must run before tools can access storage."
        )
    return _registry


def set_registry(registry: StorageRegistry | None) -> None:
    """Install (or clear, with ``None``) the process-wide registry."""
    global _registry
    _registry = registry


_settings: Settings | None = None


def get_settings() -> Settings:
    if _settings is None:
        raise RuntimeError(
            "Settings not initialised — build_app() must run first."
        )
    return _settings


def set_settings(settings: Settings | None) -> None:
    global _settings
    _settings = settings


def build_download_url(uri: str, user_id: str) -> str | None:
    """Return a ``/files/{token}`` URL if the server is configured for it.

    Requires both ``signing_secret`` and ``download_base_url`` in
    :class:`Settings`; returns ``None`` otherwise so tools can silently
    omit the ``download_url`` field when the operator has not enabled it.
    """
    settings = _settings
    if settings is None or not settings.signing_secret or not settings.download_base_url:
        return None
    token = signing.sign(
        user_id=user_id,
        uri=uri,
        secret=settings.signing_secret,
        ttl_seconds=settings.download_url_ttl_seconds,
    )
    return f"{settings.download_base_url.rstrip('/')}/files/{token}"
