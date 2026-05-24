# office-mcp

MCP server that creates and edits Word, Excel, and PowerPoint
documents on the user's Microsoft OneDrive. Speaks the
**streamable-http** transport — compatible with any MCP client (e.g.
LibreChat, Claude Desktop). The storage layer is pluggable
(StorageBackend ABC); additional backends (SharePoint, S3, Azure Blob,
local volume) can be enabled when needed.

> See [CLAUDE.md](./CLAUDE.md) for the full plan and rationale.

## Tools

**17 tools** across four groups, exposed over streamable-http at
`/mcp`, plus **3 formatting-rules resources** (`office://rules/{docx,xlsx,pptx}`).
Every mutating tool returns a signed `download_url`, so generated
files are one click away from the user — no upload-back dance.

| Group | Tools |
|---|---|
| Word  | `docx_create` / `docx_read` / `docx_append` / `docx_replace` / `docx_remove` / `docx_to_pdf` |
| Excel | `xlsx_create` / `xlsx_read_range` / `xlsx_update` |
| PowerPoint | `pptx_create` / `pptx_read` / `pptx_append` / `pptx_replace` |
| Filesystem | `fs_list` / `fs_stat` / `fs_copy` / `fs_delete` |

## Quickstart (Docker)

```bash
cp .env.example .env
# Fill in OFFICE_MCP_TOKEN, OFFICE_MCP_SIGNING_SECRET (any 32+ byte value),
# and the AZURE_* credentials if you want OneDrive support.

docker compose up --build
# → office-mcp on http://localhost:18080
# → MCP endpoint at http://localhost:18080/mcp
# → health at  http://localhost:18080/health
```

## Local development

```bash
poetry install
poetry run office-mcp                 # starts uvicorn on :8080
poetry run pytest                     # full test suite
poetry run ruff check . && poetry run mypy src
```

LibreOffice is required for the xlsx formula recalc and `docx_to_pdf`
code paths — tests that need it skip themselves when it isn't on
`PATH`.

## Wiring into an MCP client

Generic MCP client config (e.g. `mcp.json`, `claude_desktop_config.json`):

```jsonc
{
  "mcpServers": {
    "office": {
      "type": "streamable-http",
      "url": "http://localhost:18080/mcp",
      "headers": {
        "Authorization": "Bearer <OFFICE_MCP_TOKEN>",
        "X-User-Authorization": "Bearer <microsoft-graph-access-token>"
      }
    }
  }
}
```

The user-level bearer must carry `aud=Graph` and the `Files.ReadWrite`
scope — that's how office-mcp acts on the user's OneDrive without an
On-Behalf-Of exchange. Clients that do SSO with Azure AD (e.g.
LibreChat with `OPENID_REUSE_TOKENS=true`) can substitute the token at
request time.

## Repo layout

```
src/office_mcp/
  server.py            FastMCP app + lifecycle + /files/{token}
  runtime.py           Shared FastMCP + registry + settings singletons
  config.py            pydantic-settings
  context.py           per-request UserContext (extracted from headers)
  storage/             StorageBackend ABC + LocalBackend (+ Graph/S3/AzBlob in V2)
  engines/             python-docx, openpyxl, python-pptx wrappers
  tools/               docx / xlsx / pptx / fs MCP tools
  resources/           office://rules/{docx,xlsx,pptx} + content/*.md
  utils/               libreoffice headless, signing, errors
docs/                  architecture notes
scripts/               demo scenarios + ops helpers
tests/                 unit / integration / e2e
```

## License

MIT — see [LICENSE](./LICENSE).
