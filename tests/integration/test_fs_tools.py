"""Integration tests for the filesystem-level tools.

Cross-backend ``fs_copy`` is exercised against an in-memory backend
registered alongside ``LocalBackend`` — we don't wait for Phases 7-8 to
validate the pattern.
"""

from __future__ import annotations

import hashlib
from collections.abc import AsyncIterator, Iterator
from datetime import UTC, datetime
from pathlib import Path

import pytest

from office_mcp.context import UserContext, bind_user
from office_mcp.runtime import set_registry
from office_mcp.storage import FileMetadata, LocalBackend, StorageBackend, StorageRegistry
from office_mcp.tools.fs_tools import fs_copy, fs_delete, fs_list, fs_stat
from office_mcp.utils.errors import StorageNotFoundError


class MemoryBackend(StorageBackend):
    """Minimal ``memory://`` backend used to validate cross-backend ops."""

    scheme = "memory"

    def __init__(self) -> None:
        # keyed by (user_id, uri) so scoping mirrors LocalBackend
        self._store: dict[tuple[str, str], bytes] = {}

    @staticmethod
    def _now() -> str:
        return datetime.now(tz=UTC).isoformat()

    def _meta(self, uri: str, data: bytes) -> FileMetadata:
        return FileMetadata(
            uri=uri,
            size_bytes=len(data),
            modified_at=self._now(),
            etag=hashlib.sha256(data).hexdigest(),
            mime_type=None,
        )

    async def read(self, uri: str, *, user_ctx: UserContext) -> bytes:
        key = (user_ctx.user_id, uri)
        if key not in self._store:
            raise StorageNotFoundError(uri)
        return self._store[key]

    async def write(
        self,
        uri: str,
        data: bytes,
        *,
        user_ctx: UserContext,
        if_match: str | None = None,
    ) -> FileMetadata:
        self._store[(user_ctx.user_id, uri)] = data
        return self._meta(uri, data)

    async def stat(self, uri: str, *, user_ctx: UserContext) -> FileMetadata:
        key = (user_ctx.user_id, uri)
        if key not in self._store:
            raise StorageNotFoundError(uri)
        return self._meta(uri, self._store[key])

    async def list(
        self, uri_prefix: str, *, user_ctx: UserContext
    ) -> AsyncIterator[FileMetadata]:
        for (uid, uri), data in sorted(self._store.items()):
            if uid == user_ctx.user_id and uri.startswith(uri_prefix.rstrip("/")):
                yield self._meta(uri, data)

    async def delete(self, uri: str, *, user_ctx: UserContext) -> None:
        key = (user_ctx.user_id, uri)
        if key not in self._store:
            raise StorageNotFoundError(uri)
        del self._store[key]


@pytest.fixture
def registry(tmp_path: Path) -> Iterator[StorageRegistry]:
    reg = StorageRegistry()
    reg.register(LocalBackend(root=tmp_path))
    reg.register(MemoryBackend())
    set_registry(reg)
    yield reg


@pytest.fixture
def yacine() -> Iterator[UserContext]:
    ctx = UserContext(user_id="yacine", email="yacine@netwatch.fr")
    with bind_user(ctx):
        yield ctx


@pytest.mark.asyncio
async def test_fs_list_enumerates_user_tree(
    registry: StorageRegistry, tmp_path: Path, yacine: UserContext
) -> None:
    # seed two files directly through the backend for speed
    local = registry.get("local://")
    await local.write("local://a.txt", b"one", user_ctx=yacine)
    await local.write("local://nested/b.txt", b"two-two", user_ctx=yacine)

    result = await fs_list("local://")
    uris = sorted(f["uri"] for f in result["files"])
    assert uris == ["local://a.txt", "local://nested/b.txt"]
    assert result["count"] == 2
    assert result["prefix"] == "local://"


@pytest.mark.asyncio
async def test_fs_stat_returns_metadata(
    registry: StorageRegistry, yacine: UserContext
) -> None:
    local = registry.get("local://")
    await local.write("local://x.txt", b"hello", user_ctx=yacine)

    meta = await fs_stat("local://x.txt")
    assert meta["file_uri"] == "local://x.txt"
    assert meta["size_bytes"] == 5
    assert meta["etag"] is not None


@pytest.mark.asyncio
async def test_fs_stat_missing_raises(
    registry: StorageRegistry, yacine: UserContext
) -> None:
    with pytest.raises(StorageNotFoundError):
        await fs_stat("local://ghost.txt")


@pytest.mark.asyncio
async def test_fs_copy_within_local_backend(
    registry: StorageRegistry, tmp_path: Path, yacine: UserContext
) -> None:
    local = registry.get("local://")
    await local.write("local://src.txt", b"payload", user_ctx=yacine)

    result = await fs_copy("local://src.txt", "local://dst.txt")
    assert result["file_uri"] == "local://dst.txt"
    assert result["source_uri"] == "local://src.txt"
    assert result["size_bytes"] == 7

    copied = await local.read("local://dst.txt", user_ctx=yacine)
    assert copied == b"payload"


@pytest.mark.asyncio
async def test_fs_copy_across_backends(
    registry: StorageRegistry, yacine: UserContext
) -> None:
    local = registry.get("local://")
    memory = registry.get("memory://")
    await local.write("local://source.bin", b"cross-backend", user_ctx=yacine)

    result = await fs_copy("local://source.bin", "memory://dest.bin")
    assert result["file_uri"] == "memory://dest.bin"
    assert (await memory.read("memory://dest.bin", user_ctx=yacine)) == b"cross-backend"


@pytest.mark.asyncio
async def test_fs_delete_removes_file(
    registry: StorageRegistry, yacine: UserContext
) -> None:
    local = registry.get("local://")
    await local.write("local://tmp.txt", b"bye", user_ctx=yacine)

    await fs_delete("local://tmp.txt")
    with pytest.raises(StorageNotFoundError):
        await fs_stat("local://tmp.txt")


@pytest.mark.asyncio
async def test_fs_copy_unknown_source_scheme_rejected(
    registry: StorageRegistry, yacine: UserContext
) -> None:
    from office_mcp.utils.errors import UnknownSchemeError

    with pytest.raises(UnknownSchemeError):
        await fs_copy("s3://nope/foo", "local://foo")


@pytest.mark.asyncio
async def test_fs_tools_require_user_context(registry: StorageRegistry) -> None:
    with pytest.raises(RuntimeError, match="No UserContext bound"):
        await fs_list("local://")
