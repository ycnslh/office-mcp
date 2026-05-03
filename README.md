# office-mcp

MCP server that creates and edits Word, Excel, and PowerPoint
documents. Designed for **LibreChat 0.8.4** over the `streamable-http`
transport, with a pluggable storage backend (local volume today;
SharePoint/OneDrive via Microsoft Graph, S3, and Azure Blob are
pending client validation).

> See [CLAUDE.md](./CLAUDE.md) for the full plan and rationale, and
> [docs/librechat-integration.md](./docs/librechat-integration.md) for
> the step-by-step integration guide.

## Status — Phase 9 (LibreChat integration)

**17 tools** across four groups (docx, xlsx, pptx, fs) exposed over
streamable-http at `/mcp`, plus **3 formatting-rules resources**
(`office://rules/{docx,xlsx,pptx}`). Every mutating tool returns a
signed `download_url`, so files generated in chat are one click away
from the user — no upload-back dance, no volume sharing with
LibreChat's storage.

| Group | Tools |
|---|---|
| Word  | `docx_create` / `docx_read` / `docx_append` / `docx_replace` / `docx_remove` / `docx_to_pdf` |
| Excel | `xlsx_create` / `xlsx_read_range` / `xlsx_write_range` / `xlsx_apply_formula` |
| PowerPoint | `pptx_create` / `pptx_add_slide` / `pptx_read` |
| Filesystem | `fs_list` / `fs_stat` / `fs_copy` / `fs_delete` |

## Quickstart (dev stack — MCP + LibreChat 0.8.4 side by side)

```bash
# 1. Configure
cp .env.example .env
# Generate the four LibreChat secrets with `openssl rand -hex 32`
# and set OFFICE_MCP_TOKEN, OFFICE_MCP_SIGNING_SECRET (can be the same),
# MINIMAX_API_KEY, and the LibreChat secrets.

# 2. Bring the stack up
docker compose -f docker-compose.yml -f docker-compose.librechat.yml up --build

#   → office-mcp  : http://localhost:18080
#   → LibreChat   : http://localhost:3080  (register a user, start chatting)
```

In the chat, tools are discoverable via the **Serveurs MCP** menu as
"Office Documents". Try the demo prompts in
[`scripts/demo_scenarios.md`](./scripts/demo_scenarios.md).

## Integrating into an existing LibreChat deployment

Copy the two drop-in files into the client's LibreChat project:

- [`docker-compose.librechat.override.yml.example`](./docker-compose.librechat.override.yml.example)
- [`librechat.yaml.example`](./librechat.yaml.example)

Full walkthrough in
[`docs/librechat-integration.md`](./docs/librechat-integration.md),
including the known LibreChat 0.8.4 header-substitution bug and the
signed-URL download flow that works around it.

## Local development without Docker

```bash
poetry install
poetry run office-mcp                 # starts uvicorn on :8080
poetry run pytest                     # runs the full test suite
poetry run ruff check . && poetry run mypy src
```

LibreOffice is required for the xlsx formula recalc and `docx_to_pdf`
code paths — tests that need it skip themselves when it isn't on
`PATH`.

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
docs/                  architecture, librechat-integration
scripts/               demo_scenarios.md
tests/                 unit / integration / e2e
```

## License

MIT — see [LICENSE](./LICENSE).
