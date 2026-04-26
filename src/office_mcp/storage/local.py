"""Local filesystem backend with per-user scoping and path-traversal guards."""

from __future__ import annotations

import hashlib
import os
from collections.abc import AsyncIterator
from datetime import UTC, datetime
from pathlib import Path

from office_mcp.context import UserContext
from office_mcp.storage.base import FileMetadata, StorageBackend
from office_mcp.utils.errors import (
    InvalidUriError,
    PathTraversalError,
    PreconditionFailedError,
    StorageNotFoundError,
)

_SCHEME = "local"
_SAFE_USER_ID_CHARS = frozenset("-_.")
_MIME_BY_EXT: dict[str, str] = {
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ".pdf": "application/pdf",
    ".txt": "text/plain",
    ".md": "text/markdown",
    ".json": "application/json",
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
}


class LocalBackend(StorageBackend):
    """Stores files under ``<root>/<user_id>/...`` on the local filesystem."""

    scheme = _SCHEME

    def __init__(self, root: Path) -> None:
        self._root = root.resolve()
        self._root.mkdir(parents=True, exist_ok=True)

    # ── path resolution ──────────────────────────────────────────────

    @staticmethod
    def _check_user_id(user_id: str) -> None:
        # Must start with an alphanumeric (blocks "." / ".." / ".hidden") and
        # then only contain alphanumerics or a small whitelist of separators.
        if (
            not user_id
            or not user_id[0].isalnum()
            or not all(c.isalnum() or c in _SAFE_USER_ID_CHARS for c in user_id)
        ):
            raise PathTraversalError(f"Unsafe user_id: {user_id!r}")

    def _user_root(self, user_ctx: UserContext) -> Path:
        self._check_user_id(user_ctx.user_id)
        path = (self._root / user_ctx.user_id).resolve()
        if not path.is_relative_to(self._root):
            raise PathTraversalError(f"user root escapes storage root: {path}")
        path.mkdir(parents=True, exist_ok=True)
        return path

    def _relative_from_uri(self, uri: str) -> str:
        if not uri.startswith(f"{_SCHEME}://"):
            raise InvalidUriError(f"expected {_SCHEME}:// URI, got {uri!r}")
        rel = uri.removeprefix(f"{_SCHEME}://").lstrip("/")
        if not rel or "\x00" in rel:
            raise InvalidUriError(f"invalid URI path: {uri!r}")
        if any(part == ".." for part in Path(rel).parts):
            raise PathTraversalError(f"URI contains '..': {uri!r}")
        return rel

    def _resolve(self, uri: str, user_ctx: UserContext) -> Path:
        rel = self._relative_from_uri(uri)
        user_root = self._user_root(user_ctx)
        target = (user_root / rel).resolve()
        if not target.is_relative_to(user_root):
            raise PathTraversalError(f"URI escapes user scope: {uri!r}")
        return target

    # ── metadata helpers ─────────────────────────────────────────────

    @staticmethod
    def _etag(data: bytes) -> str:
        return hashlib.sha256(data).hexdigest()

    @staticmethod
    def _mtime_iso(path: Path) -> str:
        return datetime.fromtimestamp(path.stat().st_mtime, tz=UTC).isoformat()

    @staticmethod
    def _mime(path: Path) -> str | None:
        return _MIME_BY_EXT.get(path.suffix.lower())

    def _metadata(self, uri: str, path: Path, *, data: bytes | None = None) -> FileMetadata:
        payload = data if data is not None else path.read_bytes()
        return FileMetadata(
            uri=uri,
            size_bytes=path.stat().st_size,
            modified_at=self._mtime_iso(path),
            etag=self._etag(payload),
            mime_type=self._mime(path),
        )

    # ── StorageBackend API ───────────────────────────────────────────

    async def read(self, uri: str, *, user_ctx: UserContext) -> bytes:
        target = self._resolve(uri, user_ctx)
        if not target.is_file():
            raise StorageNotFoundError(uri)
        return target.read_bytes()

    async def write(
        self,
        uri: str,
        data: bytes,
        *,
        user_ctx: UserContext,
        if_match: str | None = None,
    ) -> FileMetadata:
        target = self._resolve(uri, user_ctx)
        if if_match is not None:
            if not target.is_file():
                raise PreconditionFailedError(
                    f"if_match supplied but {uri} does not exist"
                )
            current = self._etag(target.read_bytes())
            if current != if_match:
                raise PreconditionFailedError(
                    f"ETag mismatch for {uri}: have {current}, expected {if_match}"
                )
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(data)
        return self._metadata(uri, target, data=data)

    async def stat(self, uri: str, *, user_ctx: UserContext) -> FileMetadata:
        target = self._resolve(uri, user_ctx)
        if not target.is_file():
            raise StorageNotFoundError(uri)
        return self._metadata(uri, target)

    async def list(
        self,
        uri_prefix: str,
        *,
        user_ctx: UserContext,
    ) -> AsyncIterator[FileMetadata]:
        user_root = self._user_root(user_ctx)
        if uri_prefix in (f"{_SCHEME}://", f"{_SCHEME}:///"):
            base_path = user_root
        else:
            if not uri_prefix.startswith(f"{_SCHEME}://"):
                raise InvalidUriError(f"expected {_SCHEME}:// URI, got {uri_prefix!r}")
            rel = uri_prefix.removeprefix(f"{_SCHEME}://").strip("/")
            if any(part == ".." for part in Path(rel).parts):
                raise PathTraversalError(f"URI contains '..': {uri_prefix!r}")
            base_path = (user_root / rel).resolve()
            if not base_path.is_relative_to(user_root):
                raise PathTraversalError(f"URI escapes user scope: {uri_prefix!r}")
        if not base_path.is_dir():
            return
        for path in sorted(base_path.rglob("*")):
            if path.is_file():
                rel_path = path.relative_to(user_root).as_posix()
                yield self._metadata(f"{_SCHEME}://{rel_path}", path)

    async def delete(self, uri: str, *, user_ctx: UserContext) -> None:
        target = self._resolve(uri, user_ctx)
        if not target.is_file():
            raise StorageNotFoundError(uri)
        target.unlink()

    async def health_check(self) -> str:
        if not self._root.is_dir():
            raise RuntimeError(f"root {self._root} is not a directory")
        if not os.access(self._root, os.W_OK):
            raise RuntimeError(f"root {self._root} is not writable")
        return "ok"
