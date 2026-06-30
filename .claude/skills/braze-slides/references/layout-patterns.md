# Layout patterns

Map each piece of source content to the closest archetype below. All are methods
on `BrazeDeck` (see `scripts/braze_deck.py`). Author in "design inches"
(13.333 × 7.5); the builder scales to Braze's 10 × 5.625in.

Every archetype draws the leadline (purple title + gold sub-line + rule) and the
logo for you. Pass `page=N` to add a page number.

## title(title, kicker="AI Decisioning", subtitle=None, confidential=True)
Opening slide. Big purple title, gold rule, optional subtitle, "Confidential".

## section(title, subtitle=None, kicker=None)
Full-bleed lavender divider between parts of the deck. Large purple headline.

## two_column(title, left, right, sub, kicker, note, page)
Two white cards side by side, each `(heading, [bullets])`; gold bullet dots.
`note` = a closing line (string, or a list of run-dicts for mixed styling).
Use for: at-launch/ongoing, measured/billed, before/after, pros/cons.

## cards(title, items, sub, kicker, anchor, page)
A row of numbered cards. `items = [(number, heading, description), ...]`
(2–4 works best). `anchor=(lead, rest)` adds a lavender summary bar beneath.
Use for: components, steps, pillars, options.

## stat_row(title, stats, sub, kicker, page)  — or as a band on an existing slide
A row of lavender stat chips. `stats = [(value, label), ...]`.
Call standalone, or pass `slide=s, y=...` to drop the band onto a bespoke slide.
Use for: proof points, KPIs, headline metrics.

## table(title, headers, rows, sub, kicker, note, col_widths, cell_color, page)
Braze table: purple header row, white/lavender zebra rows, hairline grid.
`rows` = list of string-lists. `col_widths` in design inches (sum ≈ 12.133).
`cell_color(ri, ci, value) -> RGBColor|None` tints/bolds specific cells — e.g.
status columns (Completed→purple, In flight→gold, Upcoming→mute). Keep tables to
~12 rows; shrink `body_size`/`row_h` if dense.

## bar_chart(title, cats, values, top_labels, baseline, baseline_label, subcats, note)
Vertical bars drawn as native shapes (full brand control — never paste a chart
image). `baseline` draws a gold reference line and colours bars below it
periwinkle, at/above it purple. `top_labels` sit above bars; `cats`+`subcats`
below. Use for: ramps, month-by-month value, comparisons.

## callout(slide, l, t, w, h, lead, rest)
A purple emphasis banner with a bold gold lead-in + white body. Use for "the ask"
or a key takeaway. Call on any slide.

## closing(headline, subline, footer)
Lavender end card with purple headline, gold rule, logo, footer line.

## Bespoke slides
When nothing fits, start from `d.blank()`, call `d.header(s, title, sub=, kicker=,
page=)` to get the leadline + content-top y, then compose with primitives:
`d.box/d.para` (text), `d.card` (rounded panel), `d.rect`, `d.divider`,
`d.accent_bar`, `d.oval`, `d.footnote`. Stay on the grid (margin 0.6, content
width 12.133) and use only palette colours. Slides 3, 7, 9 and 10 of
`example_blc_deck.py` are bespoke and show the pattern (price hero, bar chart +
data table, formula + equation, dense table + highlights + ask).

## Picking an archetype (quick guide)
| Source content | Archetype |
|---|---|
| Cover / part divider | `title` / `section` |
| 2 lists, comparison | `two_column` |
| 3–4 components/steps/pillars | `cards` |
| KPIs / proof points | `stat_row` |
| Rows of data, schedule, plan | `table` |
| Trend / ramp / monthly value | `bar_chart` |
| The ask / key takeaway | `callout` |
| Big hero number, custom diagram | bespoke (`blank` + primitives) |
