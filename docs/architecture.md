# Architecture

> Detailed write-up lands in Phase 9. For now, the canonical reference is the
> diagram and rationale in [`../CLAUDE.md`](../CLAUDE.md).

## Quick mental model

```
LibreChat (api) ── streamable-http ──▶ office-mcp (FastMCP, :8080)
                                        │
                                        ├── Tools (docx / xlsx / pptx / fs)
                                        ├── Engines (python-docx / openpyxl / python-pptx)
                                        └── StorageBackend (Local / Graph / S3 / AzBlob)
```

Per-request `UserContext` is extracted from the `X-LibreChat-User-Id` and
`X-LibreChat-User-Email` headers injected by LibreChat. All storage operations
are scoped by `user_id`, so a user never sees another user's files.
