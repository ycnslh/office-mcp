"""Process-wide runtime state: the FastMCP instance and the storage registry.

Kept in its own module so tool/resource modules can import the ``mcp``
instance without creating a cycle through ``office_mcp.server``.
"""

from __future__ import annotations

from fastmcp import FastMCP

from office_mcp.storage import StorageRegistry

mcp: FastMCP = FastMCP("office-mcp")

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
