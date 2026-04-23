# PowerPoint Deck Formatting Rules

These rules govern `.pptx` files produced by office-mcp. The engine picks
the colors and fonts; the rules below constrain how you should structure
slides through tool calls.

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
| `blank`        | Full-bleed imagery or custom composition (rare).             |

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

## Templates

For recurring decks (QBR, product launch, hiring pitch) prefer a
template. List available templates with `office://templates/list`
when that resource ships (V2).
