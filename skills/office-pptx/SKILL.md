---
name: office-pptx
description: Use when creating or editing PowerPoint (.pptx) decks. The skill ships executable Python scripts (scripts/create.py and scripts/edit.py) that the agent runs inside the LibreChat code interpreter — the generated .pptx then appears as a rich file artifact in the chat. Encodes the theme catalog (corporate / minimal / cyberpunk), the slide-layout decision tree (title / section / content / two-content), copy rules (one idea per slide, ≤7-10 words per bullet, parallel grammar), and the batched-by-intent calling pattern.
invocationMode: auto
---

# PowerPoint Deck Formatting Rules

These rules govern `.pptx` files produced by `office_doctrine`. The
engine picks the colors and fonts; the rules below constrain how you
should structure slides through the create/edit scripts.

## Themes

Pick one theme per deck and stick with it. Available themes:

| Theme        | Use for                                                        |
|--------------|----------------------------------------------------------------|
| `corporate`  | Client-facing reports, board decks, sales pitches.             |
| `minimal`    | Typography-heavy talks, keynotes, editorial decks.             |
| `cyberpunk`  | Internal / hackathon / demo decks with personality.            |

Mixing themes across slides in the same deck breaks the visual rhythm.

## Slide layouts

| Layout         | When to use it                                               |
|----------------|--------------------------------------------------------------|
| `title`        | First slide only — deck title + subtitle / author.           |
| `section`      | Chapter divider between major parts of a long deck.          |
| `content`      | Default: a title + a bulleted body.                          |
| `two-content`  | Comparison / before-after / pros vs cons.                    |

## Copy on slides

- **One idea per slide.** If the bullet list goes past 6 items, split
  into two slides.
- **Max 7-10 words per bullet** in presenter-facing decks. Longer text is
  fine for *reading* decks (handouts) but hurts live delivery.
- **Parallel grammar** across sibling bullets (all imperatives, all noun
  phrases, etc.). Breaks the readability pattern when mixed.
- Use the slide **title** as the headline: it should stand alone without
  the body.

## Structure

- Title slide → executive summary → body → conclusion → appendix.
- Open with the "why", not the "what". First content slide should answer
  "why should I care?".
- Group related slides behind a `section` divider.
- Close with a single-line call to action or the next step.

## Visual hygiene

- Do NOT put full paragraphs into slide bodies. Put them in the speaker
  notes (V2 feature).
- Do NOT mix serif and sans-serif on the same slide.
- Do NOT animate bullet entries for a business audience — it adds
  cognitive load without value.

## Usage

The skill produces real `.pptx` files via the LibreChat code
interpreter. The scripts are available at `office-pptx/scripts/`
(the path the skill is mounted under inside the sandbox, per
[LibreChat PR #12831](https://github.com/danny-avila/LibreChat/pull/12831)).
Always invoke with `python3` — the code interpreter image does **not**
expose a `python` alias.

### Create a new deck — `scripts/create.py`

Payload schema (stdin JSON):

```json
{
  "theme": "corporate",
  "slides": [
    {"type": "title",       "title": "...", "subtitle": "..."},
    {"type": "content",     "title": "...", "bullets": ["...", "..."]},
    {"type": "two-content", "title": "...",
                            "left": ["..."], "right": ["..."]},
    {"type": "section",     "title": "..."}
  ]
}
```

`theme` defaults to `corporate`; other supported values are `minimal`
and `cyberpunk`. Every slide spec must carry a non-empty `title`;
`content` slides need non-empty `bullets`, `two-content` slides need
both non-empty `left` and `right`. **Build the entire deck in one
call** — pass every slide in the same payload.

Invocation from the code interpreter:

```python
import json, subprocess

payload = {
    "theme": "corporate",
    "slides": [
        {"type": "title",
         "title": "Pivot to Skills",
         "subtitle": "Yacine — Netwatch IA"},
        {"type": "section", "title": "Why now"},
        {"type": "content",
         "title": "Three drivers",
         "bullets": ["MCP previews dead-end",
                     "Skills + code interpreter native",
                     "Same doctrine, simpler delivery"]},
        {"type": "two-content",
         "title": "Old vs new",
         "left":  ["MCP server", "Custom rendering", "OneDrive round-trips"],
         "right": ["Self-hosted CI", "Native preview", "In-chat artifact"]},
    ],
}
result = subprocess.run(
    ["python3", "office-pptx/scripts/create.py"],
    input=json.dumps(payload).encode(),
    capture_output=True,
    check=True,
)
with open("pivot.pptx", "wb") as f:
    f.write(result.stdout)
```

### Edit an existing deck — `scripts/edit.py`

Same stdin-JSON pattern, two actions:

```json
{
  "action": "append",
  "input_path": "/path/to/existing.pptx",
  "params": {
    "slides": [{"type": "section", "title": "Appendix"}],
    "theme":  "corporate"
  }
}
```

```json
{
  "action": "replace",
  "input_path": "/path/to/existing.pptx",
  "params": {
    "updates": [
      {"slide_index": 0,
       "title": "New Title",
       "bullets": ["override line 1", "override line 2"]}
    ],
    "theme": "corporate"
  }
}
```

`theme` is optional in both cases and defaults to `corporate`. For
`replace`, each update needs a `slide_index` (0-based); any of
`title`, `bullets`, `left`, `right` may be set to overwrite that part
of the slide.

```python
payload = {
    "action": "append",
    "input_path": "pivot.pptx",
    "params": {
        "slides": [{"type": "content",
                    "title": "Next steps",
                    "bullets": ["ADR + tag v0.1.0",
                                "Archive office-mcp",
                                "Publish article"]}],
        "theme": "corporate",
    },
}
result = subprocess.run(
    ["python3", "office-pptx/scripts/edit.py"],
    input=json.dumps(payload).encode(),
    capture_output=True,
    check=True,
)
with open("pivot.pptx", "wb") as f:
    f.write(result.stdout)
```

Always use `python3`, never `python` — there is no symlink in the
code interpreter image.
