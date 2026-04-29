"""Per-URI :class:`asyncio.Lock` registry to serialise concurrent mutations.

Why this exists
---------------
Mutating a Graph-backed file goes read → modify → write. When an LLM
fires several pptx_set_*/pptx_add_*_slide calls concurrently against
the same file, two reads can land before either write does — and the
second write triggers a Graph 409 (``resourceModified``) because the
underlying eTag advanced after the first write.

The fix is to hold a per-URI lock around the whole read-modify-write so
the second call waits for the first to finish before starting. We do
this in-process: this MCP server has a single Python process; if we
ever scale out, the same protection has to come from Graph's own
``If-Match`` semantics rather than this map.

Memory note
-----------
The lock dict grows with the number of distinct URIs ever touched (no
eviction). For office-mcp's expected load — a handful of decks per
user per session — the leak is negligible. If we ever lift this past
"demo-grade", swap to a TTL cache.
"""

from __future__ import annotations

import asyncio
from threading import Lock as _ThreadLock

#: ``defaultdict`` would race because two coroutines hitting it
#: simultaneously can both observe a missing key and both create a
#: fresh Lock — our caller would then think they hold the same lock
#: but actually hold two distinct ones. Guard the read-or-create with
#: a tiny stdlib lock instead.
_LOCKS: dict[str, asyncio.Lock] = {}
_REGISTRY_LOCK = _ThreadLock()


def for_uri(uri: str) -> asyncio.Lock:
    """Return the :class:`asyncio.Lock` associated with ``uri``.

    The same lock is returned for the same ``uri`` across calls, so
    callers nesting ``async with for_uri(uri):`` blocks on the same
    file are mutually exclusive.
    """
    with _REGISTRY_LOCK:
        lock = _LOCKS.get(uri)
        if lock is None:
            lock = asyncio.Lock()
            _LOCKS[uri] = lock
        return lock
