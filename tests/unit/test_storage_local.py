"""Unit tests for the local filesystem storage backend."""

from __future__ import annotations

from pathlib import Path

import pytest

from office_mcp.context import UserContext
from office_mcp.storage import LocalBackend
from office_mcp.utils.errors import (
    InvalidUriError,
    PathTraversalError,
    PreconditionFailedError,
    StorageNotFoundError,
)


@pytest.fixture
def backend(tmp_path: Path) -> LocalBackend:
    return LocalBackend(root=tmp_path)


@pytest.fixture
def alice() -> UserContext:
    return UserContext(user_id="alice", email="alice@example.com")


@pytest.fixture
def bob() -> UserContext:
    return UserContext(user_id="bob")


@pytest.mark.asyncio
async def test_write_then_read_roundtrip(
    backend: LocalBackend, alice: UserContext
) -> None:
    meta = await backend.write("local://reports/q3.docx", b"hello", user_ctx=alice)

    assert meta.uri == "local://reports/q3.docx"
    assert meta.size_bytes == 5
    assert meta.etag is not None and len(meta.etag) == 64
    assert meta.mime_type == (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

    assert await backend.read("local://reports/q3.docx", user_ctx=alice) == b"hello"


@pytest.mark.asyncio
async def test_users_are_isolated(
    backend: LocalBackend, alice: UserContext, bob: UserContext
) -> None:
    await backend.write("local://shared.txt", b"alice-only", user_ctx=alice)

    with pytest.raises(StorageNotFoundError):
        await backend.read("local://shared.txt", user_ctx=bob)


@pytest.mark.asyncio
async def test_user_directories_are_separate_on_disk(
    backend: LocalBackend, tmp_path: Path, alice: UserContext, bob: UserContext
) -> None:
    await backend.write("local://doc.txt", b"a", user_ctx=alice)
    await backend.write("local://doc.txt", b"b", user_ctx=bob)

    assert (tmp_path / "alice" / "doc.txt").read_bytes() == b"a"
    assert (tmp_path / "bob" / "doc.txt").read_bytes() == b"b"


@pytest.mark.asyncio
@pytest.mark.parametrize(
    "uri",
    [
        "local://../etc/passwd",
        "local://reports/../../etc/passwd",
        "local://a/./b/../../../outside.txt",
    ],
)
async def test_path_traversal_rejected(
    backend: LocalBackend, alice: UserContext, uri: str
) -> None:
    with pytest.raises(PathTraversalError):
        await backend.read(uri, user_ctx=alice)


@pytest.mark.asyncio
@pytest.mark.parametrize(
    "uri",
    [
        "local://",
        "local://\x00nullbyte",
        "sp://other-scheme/file.docx",
        "not-a-uri",
    ],
)
async def test_invalid_uris_rejected(
    backend: LocalBackend, alice: UserContext, uri: str
) -> None:
    with pytest.raises(InvalidUriError):
        await backend.read(uri, user_ctx=alice)


@pytest.mark.asyncio
@pytest.mark.parametrize(
    "user_id", ["../escape", ".", "..", ".hidden", "", "has space", "slash/in/id"]
)
async def test_unsafe_user_id_rejected(backend: LocalBackend, user_id: str) -> None:
    with pytest.raises(PathTraversalError):
        await backend.write(
            "local://x.txt", b"x", user_ctx=UserContext(user_id=user_id)
        )


@pytest.mark.asyncio
async def test_etag_mismatch_blocks_overwrite(
    backend: LocalBackend, alice: UserContext
) -> None:
    first = await backend.write("local://doc.txt", b"v1", user_ctx=alice)
    assert first.etag is not None

    with pytest.raises(PreconditionFailedError):
        await backend.write(
            "local://doc.txt", b"v2", user_ctx=alice, if_match="0" * 64
        )

    overwritten = await backend.write(
        "local://doc.txt", b"v2", user_ctx=alice, if_match=first.etag
    )
    assert overwritten.etag != first.etag
    assert await backend.read("local://doc.txt", user_ctx=alice) == b"v2"


@pytest.mark.asyncio
async def test_if_match_without_existing_file_fails(
    backend: LocalBackend, alice: UserContext
) -> None:
    with pytest.raises(PreconditionFailedError):
        await backend.write(
            "local://new.txt", b"v1", user_ctx=alice, if_match="deadbeef"
        )


@pytest.mark.asyncio
async def test_stat_returns_size_and_etag(
    backend: LocalBackend, alice: UserContext
) -> None:
    await backend.write("local://a.txt", b"hello", user_ctx=alice)

    meta = await backend.stat("local://a.txt", user_ctx=alice)
    assert meta.size_bytes == 5
    assert meta.etag is not None


@pytest.mark.asyncio
async def test_stat_missing_raises(backend: LocalBackend, alice: UserContext) -> None:
    with pytest.raises(StorageNotFoundError):
        await backend.stat("local://ghost.txt", user_ctx=alice)


@pytest.mark.asyncio
async def test_list_yields_files_in_user_scope(
    backend: LocalBackend, alice: UserContext, bob: UserContext
) -> None:
    await backend.write("local://a.txt", b"1", user_ctx=alice)
    await backend.write("local://nested/b.txt", b"22", user_ctx=alice)
    await backend.write("local://bobs-secret.txt", b"x", user_ctx=bob)

    uris = [m.uri async for m in backend.list("local://", user_ctx=alice)]
    assert uris == ["local://a.txt", "local://nested/b.txt"]


@pytest.mark.asyncio
async def test_list_with_prefix_limits_scope(
    backend: LocalBackend, alice: UserContext
) -> None:
    await backend.write("local://a.txt", b"1", user_ctx=alice)
    await backend.write("local://nested/b.txt", b"22", user_ctx=alice)

    uris = [m.uri async for m in backend.list("local://nested/", user_ctx=alice)]
    assert uris == ["local://nested/b.txt"]


@pytest.mark.asyncio
async def test_list_missing_prefix_yields_nothing(
    backend: LocalBackend, alice: UserContext
) -> None:
    uris = [m.uri async for m in backend.list("local://nope/", user_ctx=alice)]
    assert uris == []


@pytest.mark.asyncio
async def test_delete_removes_file(backend: LocalBackend, alice: UserContext) -> None:
    await backend.write("local://x.txt", b"x", user_ctx=alice)
    await backend.delete("local://x.txt", user_ctx=alice)

    with pytest.raises(StorageNotFoundError):
        await backend.read("local://x.txt", user_ctx=alice)


@pytest.mark.asyncio
async def test_delete_missing_raises(
    backend: LocalBackend, alice: UserContext
) -> None:
    with pytest.raises(StorageNotFoundError):
        await backend.delete("local://ghost.txt", user_ctx=alice)
