"""Scheme → backend routing."""

from __future__ import annotations

from office_mcp.storage.base import StorageBackend
from office_mcp.utils.errors import InvalidUriError, UnknownSchemeError


class StorageRegistry:
    """Looks up the storage backend that owns a given URI."""

    def __init__(self) -> None:
        self._backends: dict[str, StorageBackend] = {}

    def register(self, backend: StorageBackend) -> None:
        self._backends[backend.scheme] = backend

    def get(self, uri: str) -> StorageBackend:
        if "://" not in uri:
            raise InvalidUriError(
                f"URI {uri!r} is missing a scheme. "
                f"Available schemes: {self._available_schemes()}."
            )
        scheme = uri.split("://", 1)[0]
        backend = self._backends.get(scheme)
        if backend is None:
            raise UnknownSchemeError(
                f"URI scheme {scheme!r} has no registered backend. "
                f"Available schemes: {self._available_schemes()}. "
                f"Use {self._suggest()} instead."
            )
        return backend

    def _available_schemes(self) -> str:
        return ", ".join(sorted(f"{s}://" for s in self._backends)) or "(none)"

    def _suggest(self) -> str:
        if not self._backends:
            return "(no backend registered)"
        return f"{next(iter(sorted(self._backends)))}://your/path/here"
