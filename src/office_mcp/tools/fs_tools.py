"""MCP tools that sit directly on top of the :class:`StorageBackend` API.

These are format-agnostic: they let the LLM inspect, move, or delete
files without needing the engine layer. Cross-backend operations work
transparently — ``fs_copy`` reads bytes from the source backend and
writes them to the destination, so other backends (S3, Azure Blob,
SharePoint) will plug in without changes once they land.
"""

from __future__ import annotations

from typing import Any

from office_mcp.context import require_current_user
from office_mcp.instrumentation import register_tool
from office_mcp.runtime import download_url_for, get_registry


async def fs_list(uri_prefix: str) -> dict[str, Any]:
    """List files in a OneDrive folder.

    Call this only when the user explicitly asks to see folder contents,
    or when you need to find an existing file to edit. Do NOT use it as
    a "scope" or "check existence" probe before a ``*_create`` call —
    create tools handle auth and overwrite on their own, so the list
    pre-call only adds latency.

    The bare ``od://`` lists the drive root; ``od://AI Generated/``
    lists the folder office-mcp drops generated files in by default.
    First call against ``od://`` in a fresh chat returns
    ``{"auth_required": true, "login_url": ...}`` — surface the URL,
    wait for the user, then retry with the same args.

    Args:
        uri_prefix: A directory-like URI. Examples: ``od://``,
            ``od://AI Generated/``, ``od://Documents/``.
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
                "download_url": await download_url_for(meta, user),
            }
        )
    return {
        "prefix": uri_prefix,
        "count": len(files),
        "files": files,
    }


async def fs_stat(file_uri: str) -> dict[str, Any]:
    """Return metadata (size, mtime, etag, mime) for a single file.

    Use this when you need the size or etag of an existing file. Do NOT
    use it to "check if a file exists" before a ``*_create`` call —
    create tools overwrite by design, so the stat pre-call adds latency
    for nothing.

    ``file_uri`` is on the user's OneDrive, e.g.
    ``od://AI Generated/q3.docx``. The first ``od://`` call in a fresh
    chat may return an ``auth_required`` payload — surface the login
    URL and retry.
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
        "download_url": await download_url_for(meta, user),
    }


async def fs_copy(source_uri: str, destination_uri: str) -> dict[str, Any]:
    """Copy a file from one OneDrive URI to another.

    Use this to duplicate a document under a new name or move it to a
    different folder. Same-scheme copies (``od://`` → ``od://``) read
    the bytes through the backend and write them at the destination —
    the source is unchanged.
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
        "download_url": await download_url_for(meta, user),
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
