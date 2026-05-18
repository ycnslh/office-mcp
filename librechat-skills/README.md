# office-mcp — LibreChat Skill bundles

Three Agent Skills (LibreChat ≥ 0.8.6-rc1) that surface the office-mcp
formatting doctrine as importable bundles. Same content as the MCP
resources `office://rules/{docx,xlsx,pptx}`, but packaged so any
LibreChat agent can pull them in without going through MCP.

- `office-docx-rules/` — Word formatting doctrine
- `office-xlsx-rules/` — Excel finance-style conventions
- `office-pptx-rules/` — PowerPoint deck rules

Each bundle is a single `SKILL.md` with a YAML frontmatter
(`name`, `description`, `invocationMode: auto`) parsed by
[client/src/components/Skills/utils/parseSkillMd.ts](https://github.com/danny-avila/LibreChat/blob/main/client/src/components/Skills/utils/parseSkillMd.ts).
`invocationMode: auto` makes LibreChat surface the skill to the model
based on the description — the LLM decides whether to pull it in for a
given turn.

## Build

```sh
./build.sh
```

Produces `dist/office-{docx,xlsx,pptx}-rules.zip`. Each archive has
`SKILL.md` at the root, which matches the layout accepted by
[packages/api/src/skills/import.ts](https://github.com/danny-avila/LibreChat/blob/main/packages/api/src/skills/import.ts)
(SKILL.md at root or one level deep, ≤ 50 MB compressed, ≤ 500 entries,
each file ≤ 10 MB — easily met).

## Import into LibreChat staging

The PoC instance lives at `librechat-staging.reho.dev` (pinned to
v0.8.6-rc1, see [docker-compose.librechat-staging.yml](../docker-compose.librechat-staging.yml)).

UI path:
1. Sign in on `librechat-staging.reho.dev`
2. Open Skills (Settings → Skills, or the `$` picker in the chat input)
3. Upload each `dist/*.zip` via "Import Skill"
4. Attach each skill to an agent that has the office-mcp tools enabled

Programmatic path (when iterating):
```sh
TOKEN=...  # JWT from /api/auth/me
for z in dist/*.zip; do
  curl -X POST https://librechat-staging.reho.dev/api/skills/import \
    -H "Authorization: Bearer $TOKEN" \
    -F "file=@${z}"
done
```

## Keeping content in sync

The bodies of the three `SKILL.md` files mirror the source-of-truth
markdown at [src/office_mcp/resources/content/](../src/office_mcp/resources/content/).
If you change a rule there, mirror it in the matching skill — there is
no build pipeline tying the two together yet (intentional simplicity
for the PoC; revisit if the duplication becomes painful).
