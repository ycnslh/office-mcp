"""Process-wide runtime state: the FastMCP instance and the storage registry.

Kept in its own module so tool/resource modules can import the ``mcp``
instance without creating a cycle through ``office_mcp.server``.
"""

from __future__ import annotations

from fastmcp import FastMCP

from office_mcp.storage import StorageRegistry

_INSTRUCTIONS = """\
office-mcp creates and edits Word, Excel and PowerPoint documents.

ADDRESSING FILES
  Every tool takes a `file_uri` like `local://path/to/file.ext`.
  In this deployment, ONLY the `local://` scheme is wired — do not invent
  `s3://`, `az://`, `sp://` or `od://` URIs (Phases 7-8 will add them).
  Path components are scoped per user, so a leading `/` is unnecessary.

WORKFLOW
  1. Before creating a document, consult the matching `office://rules/<format>`
     resource (docx, xlsx) for formatting conventions.
  2. Compose the document with the granular tools (create → write/append →
     read to verify).
  3. Always echo the resulting `file_uri` to the user so they can retrieve it.

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
