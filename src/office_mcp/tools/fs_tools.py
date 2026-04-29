"""MCP tools that sit directly on top of the :class:`StorageBackend` API.

These are format-agnostic: they let the LLM inspect, move, or delete
files without needing the engine layer. Cross-backend operations work
transparently — ``fs_copy`` reads bytes from the source backend and
writes them to the destination, so ``local://`` ↔ ``s3://`` copies will
just work once those backends land in Phases 7-8.
"""

from __future__ import annotations

from typing import Any

from office_mcp.context import require_current_user
from office_mcp.instrumentation import register_tool
from office_mcp.runtime import build_download_url, get_registry


async def fs_list(uri_prefix: str) -> dict[str, Any]:
    """List files under a URI prefix.

    Pick the scheme from the user's intent — never default silently:

    * ``od://`` — the user's Microsoft OneDrive. Use whenever they say
      "OneDrive", "my Microsoft files", "cloud", "Office 365". The bare
      ``od://`` lists the drive root; ``od://Documents/`` lists a folder.
      First call returns ``{"auth_required": true, "login_url": ...}`` —
      surface the URL, wait for the user, then retry with the same args.
    * ``local://`` — the user's per-account scratch space on the server.
      Use only for drafts the user hasn't asked to put in the cloud.
      ``local://`` lists the whole scope; ``local://reports/`` a folder.

    Args:
        uri_prefix: A directory-like URI. Examples: ``od://``,
            ``od://Documents/``, ``local://``, ``local://reports/``.
    """
    user = require_current_user()
    backend = get_registry().get(uri_prefix)
    files: list[dict[str, Any]] = []
    async for meta in backend.list(uri_prefix, user_ctx=user):
        files.append(
            {
                "uri": meta.uri,
                "size_bytes": meta.size_bytes,
                "modified_at": meta.modified_at,
                "mime_type": meta.mime_type,
                "download_url": build_download_url(meta.uri, user.user_id),
            }
        )
    return {
        "prefix": uri_prefix,
        "count": len(files),
        "files": files,
    }


async def fs_stat(file_uri: str) -> dict[str, Any]:
    """Return metadata (size, mtime, etag, mime) for a single file.

    ``file_uri`` carries the storage scheme — ``od://Documents/q3.docx``
    for OneDrive, ``local://reports/q3.docx`` for the server scratch
    space. ``od://`` may return an ``auth_required`` payload on first
    call; surface the login URL and retry.
    """
    user = require_current_user()
    backend = get_registry().get(file_uri)
    meta = await backend.stat(file_uri, user_ctx=user)
    return {
        "file_uri": meta.uri,
        "size_bytes": meta.size_bytes,
        "modified_at": meta.modified_at,
        "etag": meta.etag,
        "mime_type": meta.mime_type,
        "download_url": build_download_url(meta.uri, user.user_id),
    }


async def fs_copy(source_uri: str, destination_uri: str) -> dict[str, Any]:
    """Copy a file between URIs.

    Works across storage backends: ``local://`` ↔ ``od://`` is the
    typical case (publish a draft from local scratch to the user's
    OneDrive, or pull a OneDrive file down to local for editing). Same-
    scheme copies are handled identically — no fast path, but correct.
    """
    user = require_current_user()
    source_backend = get_registry().get(source_uri)
    destination_backend = get_registry().get(destination_uri)
    data = await source_backend.read(source_uri, user_ctx=user)
    meta = await destination_backend.write(
        destination_uri, data, user_ctx=user
    )
    return {
        "source_uri": source_uri,
        "file_uri": meta.uri,
        "size_bytes": meta.size_bytes,
        "etag": meta.etag,
        "download_url": build_download_url(meta.uri, user.user_id),
        "summary": (
            f"Copied {source_uri} → {meta.uri} ({meta.size_bytes} bytes)"
        ),
    }


async def fs_delete(file_uri: str) -> dict[str, Any]:
    """Delete a file."""
    user = require_current_user()
    backend = get_registry().get(file_uri)
    await backend.delete(file_uri, user_ctx=user)
    return {
        "file_uri": file_uri,
        "summary": f"Deleted {file_uri}",
    }


# Register with FastMCP imperatively — see tools/docx_tools.py for the why.
for _fn in (fs_list, fs_stat, fs_copy, fs_delete):
    register_tool(_fn)
