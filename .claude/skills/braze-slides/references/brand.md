# Braze slide brand spec

Extracted from the official Braze "Decisioning Studio" slideware theme. These are
the values `braze_deck.py` uses. Treat them as canonical.

## Colour palette
| Token | Hex | Use |
|-------|-----|-----|
| Purple (brand) | `#801ED7` | Headers, section labels, emphasis fills, table header rows, bars |
| Deep purple | `#300266` | Strong/dark emphasis, "today"-type states |
| Gold | `#FFA524` | Sub-lines under headers, accents, dots, callout lead-ins, rules |
| Berry | `#91186E` | Rare accent only |
| Red | `#E9371F` | Warnings only — use sparingly |
| Lavender | `#EFEDFF` | Card fills, zebra rows, soft panels |
| Periwinkle | `#C9C4FF` | Secondary fills, "below-threshold" bars, light borders |
| Ink | `#1A1A24` | Body text |
| Slate | `#555266` | Secondary text |
| Mute | `#8A8699` | Captions, footnotes, page numbers |
| Hairline | `#E3DFF5` | Borders, dividers, table grid |
| On-purple | `#D9D3FB` | Light text on purple fills |
| White | `#FFFFFF` | Backgrounds, text on purple |

Backgrounds are white (content slides) or lavender `#EFEDFF` (section / closing).
No gradients, no shadows.

## Typography
- **Family:** Helvetica Neue for everything. Fallback: Arial. Nothing else.
- **Leadline (slide title):** Helvetica Neue **Bold**, purple `#801ED7`,
  ~22–24pt on a content slide, ~34–40pt on title/section.
- **Sub-line:** Bold, gold `#FFA524`, ~13pt.
- **Kicker/eyebrow:** Bold gold caps, ~11pt, letter-spaced.
- **Body:** Regular, ink/slate, ~11–12pt. **Bullets** use a gold "•".
- **Footnote:** ~7.5pt muted italic.
- Type scale (pt, design space): 40 title · 34 section · 22–24 leadline ·
  15 card heading · 13 sub-line · 11–12 body · 9 caption · 7.5 footnote.

## Logo
- Asset: `assets/brazeai_logo.png` (the full-colour BrazeAI lockup, 1200×513).
- Position: **top-right**, ~0.7in wide, ~0.34in from the top, aligned to the
  right margin. Same on every slide. Placed automatically by `header()` and the
  title/section/closing archetypes.
- On dark purple fills the colour logo won't read — keep logo slides light.

## Layout grid (design space = 13.333 × 7.5, emitted at 0.75 → 10 × 5.625in)
- Left & right margin: `0.6`. Content width: `12.133`.
- Leadline block starts at y≈`0.5`; content begins ≈`2.0` (≈`2.2` with a kicker).
- Cards: rounded corners (radius ~0.05), 1.25pt hairline border, white or
  lavender fill, optional 0.12 gold top-bar or a 0.1 left accent bar.
- Page number bottom-right; footnote along the bottom margin.

## Don'ts
- No fonts other than Helvetica Neue/Arial. No emoji, clip-art, stock photos.
- No colour outside the palette. No gradients, glows, or shadows.
- Don't centre body text or headers (headers left-align to the margin).
- Don't change the customer's numbers or wording when restyling.
