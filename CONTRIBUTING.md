# Contributing to office-mcp

Thanks for taking the time to look at the project. Small PRs are
welcome — bug fixes, new storage backends, additional formatting
rules, doc improvements.

## TL;DR

```bash
poetry install --with dev
poetry run pytest                 # full suite
poetry run ruff check .           # lint
poetry run mypy src               # type check
poetry run office-mcp             # boot on :8080
```

Conventional Commits, one logical change per PR, tests green, CI must
pass. Done.

## Prerequisites

- **Python 3.12** (`pyenv install 3.12` if you don't have it)
- **Poetry 1.8+** (`pipx install poetry` or per the
  [official installer](https://python-poetry.org/docs/#installation))
- **Docker** — only needed if you want to build/run the image or
  reproduce the CI environment locally
- **LibreOffice** (optional) — required for the `docx_to_pdf` tool
  and for xlsx formula recalculation. Tests that depend on it
  auto-skip when it's not on `PATH`.

## Setup

```bash
git clone https://github.com/soulkill3r/office-mcp
cd office-mcp
poetry install --with dev
cp .env.example .env              # fill in OFFICE_MCP_TOKEN at least
```

Run the server:

```bash
poetry run office-mcp             # binds 0.0.0.0:8080
# MCP endpoint:  http://localhost:8080/mcp
# Health:        http://localhost:8080/health
```

Or via Docker, matching the CI image:

```bash
docker compose up --build         # exposes :18080 on the host
```

## Tests

```bash
poetry run pytest                 # everything
poetry run pytest tests/unit      # fast, no LibreOffice required
poetry run pytest -k docx         # subset by keyword
poetry run pytest -v --tb=short   # verbose
```

The CI runs `ruff check . && pytest`. Anything green locally should be
green there.

## Style

- **Type hints** everywhere. `poetry run mypy src` should stay clean.
- **Ruff** for lint and import sorting. `poetry run ruff check .` and
  `poetry run ruff format .` before pushing.
- **Docstrings** on public functions, Google-style. Keep them about
  *why*, not *what* (the signature already tells you the what).
- **No comments restating the code**. Comments earn their place when
  they explain a non-obvious invariant or a subtle workaround.

## Adding a storage backend

The contract is in `src/office_mcp/storage/base.py` (`StorageBackend`
ABC). Look at `local.py` for the simplest reference and `graph.py`
for the cloud / async / retry pattern. Tests live in
`tests/unit/test_storage_*.py` and `tests/integration/`.

The `.claude/agents/storage-backend-builder.md` agent describes the
expected shape if you use Claude Code.

## Adding an MCP tool

Tools live in `src/office_mcp/tools/`. Registration pattern: write
`async def my_tool(...)` at the top of the module, then
`mcp.tool(my_tool)` at the bottom. The `@mcp.tool` decorator wraps the
function in a `FunctionTool` object that breaks direct callability
from other tool code, so we register at module bottom instead.

Tools should return a dict with at least `file_uri`, `size_bytes`,
and a `summary` string — see existing tools for the full shape
(`download_url`, `etag`, etc.).

The `.claude/agents/mcp-tool-author.md` agent codifies this convention.

## Commit messages

[Conventional Commits](https://www.conventionalcommits.org/). Used types:

- `feat:` new user-visible capability
- `fix:` bug fix
- `chore:` repo hygiene, dependency bumps
- `docs:` doc-only changes
- `test:` test-only changes
- `ci:` workflow changes
- `refactor:` no behaviour change

Subject line ≤ 72 chars. Body wraps at 72 too. Reference issues with
`Refs #123` or `Fixes #123`.

## Pull requests

- One logical change per PR. If your local branch grew several
  unrelated tweaks, split them.
- Tests must pass locally before you push.
- CI must be green before merge.
- The maintainer squash-merges PRs onto `main` (linear history).
- For non-trivial changes, open a quick issue first to align on
  approach.

## Reporting issues

- **Bug** → [GitHub Issues](https://github.com/soulkill3r/office-mcp/issues).
  Include version / commit SHA, the request that failed, expected vs
  observed, and a reproducer if possible.
- **Security vulnerability** → see [SECURITY.md](./SECURITY.md). Do not
  open a public issue.

## Architecture overview

The conventions, layering, and contracts that shape the codebase are
documented in [CLAUDE.md](./CLAUDE.md) (named that way because it
doubles as the entry-point document for [Claude Code](https://docs.claude.com/en/docs/claude-code/overview)
agents, but it's readable by humans too).
