# office-mcp — LibreChat Skill bundles

Three self-contained LibreChat Agent Skills (≥ 0.8.6-rc1) that produce
real `.docx`/`.xlsx`/`.pptx` files via the LibreChat **code interpreter**,
with no MCP server in the loop. The skills package the doctrine markdown
(formerly `office://rules/*`) together with executable Python scripts
that call the `office_doctrine` lib preinstalled in the code interpreter
image (see `docker/code-interpreter/`).

- `office-docx/` — Word: create + append/replace/remove
- `office-xlsx/` — Excel: create + update (finance color-coding, formula recalc)
- `office-pptx/` — PowerPoint: create + append/replace (themes: corporate/minimal/cyberpunk)

Each bundle has:
- `SKILL.md` (YAML frontmatter `name` / `description` / `invocationMode: auto`,
  body = formatting doctrine + `## Usage` section documenting the scripts)
- `scripts/create.py` (stdin JSON → stdout bytes)
- `scripts/edit.py` (stdin JSON with `input_path` → stdout bytes)

## Pipeline

```
user prompt
  → LLM in LibreChat agent (skill auto-selected via `description` match)
  → code interpreter: `python3 office-<fmt>/scripts/create.py < payload.json > out.<ext>`
  → `out.<ext>` written to the sandbox FS
  → LibreChat picks it up → rich artifact preview in the side panel (PR #12934)
```

The skill files are mounted inside the sandbox under `office-<fmt>/`
([LibreChat PR #12831](https://github.com/danny-avila/LibreChat/pull/12831)),
which is why scripts are invoked with that path prefix.

**Always `python3`**, never `python` — the
[LibreCodeInterpreter](https://github.com/usnavy13/LibreCodeInterpreter)
base image does not symlink `python`.

## Build

```sh
./build.sh
```

Produces `dist/office-{docx,xlsx,pptx}.zip`, each ~10–20 KB. The `office-doctrine`
wheel is **not** bundled — it lives in the code interpreter image (built by
`docker/code-interpreter/Dockerfile`). Skills stay lean.

Archive layout matches what
[packages/api/src/skills/import.ts](https://github.com/danny-avila/LibreChat/blob/main/packages/api/src/skills/import.ts)
accepts: `SKILL.md` at the root, ≤ 50 MB compressed, ≤ 500 entries.

## Import into LibreChat staging

The PoC instance lives at `librechat-staging.reho.dev` (pinned to v0.8.6-rc1,
see [docker-compose.librechat-staging.yml](../docker-compose.librechat-staging.yml)).
The code interpreter is already wired in via `LIBRECHAT_CODE_BASEURL`.

UI path:
1. Sign in on `librechat-staging.reho.dev`
2. Open Skills (Settings → Skills, or the `$` picker in the chat input)
3. Upload each `dist/*.zip` via "Import Skill"
4. Create an agent, attach the three skills, **enable Code Interpreter**
5. Prompt — e.g. "génère-moi un rapport docx de 2 paragraphes sur le projet X".
   The LLM picks `office-docx` automatically (description match), runs
   `scripts/create.py` in the sandbox, and the file appears as a rich artifact.

Programmatic path (when iterating):
```sh
TOKEN=...  # JWT from /api/auth/me
for z in dist/*.zip; do
  curl -X POST https://librechat-staging.reho.dev/api/skills/import \
    -H "Authorization: Bearer $TOKEN" \
    -F "file=@${z}"
done
```

## Keeping the doctrine in sync

The doctrine sections of the three `SKILL.md` files mirror the source-of-truth
markdown at [src/office_mcp/resources/content/](../src/office_mcp/resources/content/)
(still served by the legacy MCP server). If you change a rule there, mirror it
in the matching skill — no build pipeline ties the two together yet (intentional
simplicity for the PoC, revisit if duplication becomes painful and once the
office-mcp archival decision is made — see plan PR4).
