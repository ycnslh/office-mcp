"""Phase 0 smoke test — guarantees the package imports.

Real test suites land alongside their features in Phases 1+.
"""

from __future__ import annotations


def test_package_imports() -> None:
    import office_mcp

    assert office_mcp.__version__
