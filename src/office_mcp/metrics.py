"""Prometheus metrics — exposed via the ``/metrics`` route on the MCP server.

We keep a private ``CollectorRegistry`` rather than the global one so that
test runs don't accumulate label combinations across modules.
"""

from __future__ import annotations

from prometheus_client import CollectorRegistry, Counter, Histogram

REGISTRY = CollectorRegistry()

tool_calls = Counter(
    "office_mcp_tool_calls_total",
    "Tool invocations broken down by tool name and outcome.",
    ["tool", "status"],
    registry=REGISTRY,
)

tool_duration_seconds = Histogram(
    "office_mcp_tool_duration_seconds",
    "Wall-clock duration of MCP tool calls.",
    ["tool"],
    registry=REGISTRY,
    # Buckets sized for the actual mix: most tools <100 ms, LibreOffice
    # paths (recalc, pdf) sit around 1.5-3 s, occasional larger files.
    buckets=(0.05, 0.1, 0.25, 0.5, 1.0, 2.0, 5.0, 10.0, 30.0),
)

downloads_total = Counter(
    "office_mcp_downloads_total",
    "Signed-URL download requests by HTTP-style status.",
    ["status"],
    registry=REGISTRY,
)

rate_limited_total = Counter(
    "office_mcp_rate_limited_total",
    "MCP requests rejected by the rate limiter.",
    ["user_id"],
    registry=REGISTRY,
)
