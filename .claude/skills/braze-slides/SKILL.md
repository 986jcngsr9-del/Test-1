---
name: braze-slides
description: >
  Turn slide content — a markdown outline, rough notes, or an existing deck —
  into on-brand BrazeAI slideware as a .pptx. Use whenever the user wants slides
  made "Braze standard", on-brand, brand-compliant, or styled like Braze /
  BrazeAI / Decisioning Studio slideware; whether restyling an existing deck or
  building new bespoke slides that must look like an internal designer made them
  (not AI-generated). Produces a 10x5.625in Helvetica Neue deck with the Braze
  palette and the real BrazeAI logo.
---

# Braze slides

Generate slides that pass as professionally designed Braze slideware. The brand
system is encoded in `scripts/braze_deck.py`; you author a short Python build
script that calls it. **Do not** invent colours, fonts, or layouts — use the kit.

## When to use
- "Make these slides Braze standard / on-brand / look like Braze."
- "Restyle this deck" or "turn this outline/markdown into Braze slides."
- Building bespoke slides that still must look on-brand.

## The brand in one breath
- **Font:** Helvetica Neue everywhere. **Leadlines are Bold.** (Arial is the only acceptable fallback.)
- **Purple `#801ED7`** for headers & emphasis · **deep purple `#300266`** · **gold `#FFA524`** for sub-lines/accents · **lavender `#EFEDFF`** & **periwinkle `#C9C4FF`** for fills · black body text.
- **Every slide:** BrazeAI logo top-right (same size/position) + a leadline = bold purple title with a gold sub-line and a short gold rule.
- **Canvas:** 10 x 5.625in (Braze's real size — `braze_deck` sets this automatically).
- Full spec: `references/brand.md`. Layout catalogue: `references/layout-patterns.md`.

## Workflow
1. **Get the content.**
   - From markdown/notes: read it.
   - From an existing deck (.pptx/Google Slides): extract the *text* first
     (e.g. unzip the pptx and read `ppt/slides/*.xml`, or use a Drive
     `read_file_content`). Keep the wording verbatim unless asked to edit.
2. **Map each slide to a layout pattern** from `references/layout-patterns.md`
   (leadline/section, two-column bullets, numbered cards, stat row, table,
   bar chart, callout, closing). Pick the closest archetype; drop to a blank
   slide + primitives only when the content is genuinely bespoke.
3. **Write a build script** that imports `braze_deck` and calls the archetypes.
   `scripts/example_blc_deck.py` is a complete, copy-pasteable template that
   uses both archetypes and bespoke primitives. Start from it.
4. **Generate:** `python3 your_build.py out.pptx` (run from the `scripts/`
   folder, or add it to `sys.path`, so `import braze_deck` resolves).
5. **Verify visually.** This is the step that makes it reliable — render and
   look. In Cowork, open the .pptx (or upload to Drive → Google Slides). If a
   renderer is available locally: `soffice --headless --convert-to pdf out.pptx
   && pdftoppm -png -r 110 out.pdf page`. Check: logo present & aligned, no text
   overflow/clipping, consistent margins, leadline reads purple+gold, nothing
   off-canvas. Fix coordinates and regenerate. Iterate until it looks designed.
6. **Deliver** the .pptx. To land it in Google Slides, upload the file to Drive
   (it auto-converts) — Helvetica Neue and the layout are preserved.

## Rules that keep it from looking AI-made
- One type family (Helvetica Neue), a tight type scale, generous whitespace.
- Left-align headers to a single left margin; keep the same margin on every slide.
- Use the brand purples/gold only — no rainbow, no gradients, no drop shadows,
  no emoji, no stock clip-art. Tints come from lavender/periwinkle only.
- Tables: purple header row, white/lavender zebra rows, hairline `#E3DFF5` rules.
- Charts: draw them as native shapes via `bar_chart` (full brand control) rather
  than pasting chart images.
- Numbers/wording are content — never change them while restyling.
- Footnotes in small muted italic at the bottom; page number bottom-right.

## Files
- `scripts/braze_deck.py` — the `BrazeDeck` builder (archetypes + primitives).
- `scripts/example_blc_deck.py` — full worked example / template (11 slides).
- `assets/brazeai_logo.png` — the official BrazeAI logo (placed top-right).
- `references/brand.md` — exact palette, type scale, spacing, logo rules.
- `references/layout-patterns.md` — when/how to use each archetype.

The user may keep a folder of approved Braze example decks; if provided, open a
few and mirror their spacing/treatment. The kit already matches that slideware.
