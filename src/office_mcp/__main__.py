"""Entry point: ``python -m office_mcp`` or the ``office-mcp`` script."""

from __future__ import annotations


def main() -> None:
    """Boot the MCP server. Real implementation lands in Phase 2."""
    from office_mcp.server import run

    run()


if __name__ == "__main__":
    main()
