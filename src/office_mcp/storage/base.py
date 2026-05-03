"""Storage backend abstract base class and shared metadata type."""

from __future__ import annotations

from abc import ABC, abstractmethod
from collections.abc import AsyncIterator
from dataclasses import dataclass

from office_mcp.context import UserContext


@dataclass(frozen=True, slots=True)
class FileMetadata:
    """Lightweight descriptor returned by every backend operation."""

    uri: str
    size_bytes: int
    modified_at: str
    etag: str | None
    mime_type: str | None
    web_url: str | None = None
    """Native browser URL for the file in the backend's UI (e.g. OneDrive
    web). ``None`` for backends that don't expose one (local). Tools that
    surface a download/view link prefer this over the local signed URL
    when present, so cloud-backed files open in their native viewer."""


class StorageBackend(ABC):
    """Uniform interface over Local / Graph / S3 / Azure Blob backends.

    Every method is scoped by ``user_ctx.user_id`` — implementations MUST
    ensure a caller never observes another user's files.

    ``write`` accepts an ``if_match`` ETag for optimistic concurrency: if the
    current object ETag does not match, :class:`PreconditionFailedError` is
    raised, letting the MCP layer re-read and re-apply the change.
    """

    scheme: str

    @abstractmethod
    async def read(self, uri: str, *, user_ctx: UserContext) -> bytes:
        """Return the file contents at ``uri``."""

    @abstractmethod
    async def write(
        self,
        uri: str,
        data: bytes,
        *,
        user_ctx: UserContext,
        if_match: str | None = None,
    ) -> FileMetadata:
        """Create or replace the file at ``uri`` and return its new metadata."""

    @abstractmethod
    async def stat(self, uri: str, *, user_ctx: UserContext) -> FileMetadata:
        """Return metadata for the file at ``uri``."""

    @abstractmethod
    def list(
        self,
        uri_prefix: str,
        *,
        user_ctx: UserContext,
    ) -> AsyncIterator[FileMetadata]:
        """Yield metadata for every file under ``uri_prefix`` (recursively)."""

    @abstractmethod
    async def delete(self, uri: str, *, user_ctx: UserContext) -> None:
        """Delete the file at ``uri``. Raises if it does not exist."""

    async def health_check(self) -> str:
        """Return ``"ok"`` if the backend is reachable, raise otherwise.

        The default implementation is a no-op; backends with non-trivial
        IO (filesystem, network) should override and probe the underlying
        store. The result feeds the ``/health`` endpoint.
        """
        return "ok"
