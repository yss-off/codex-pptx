# From-Scratch Generation

Use this workflow when no suitable template exists.

## 1. Freeze the presentation contract

Confirm the topic, audience, purpose, expected duration, language, tone, delivery environment, source material, and output path. Clarify material ambiguity before polishing slides.

## 2. Plan the deck

Write one sentence describing what the audience should believe or decide after the presentation. Give every slide one job and remove slides that do not advance it.

Read the titles in order before page-level polishing: what new information does each page add, why does it follow the previous page, and what question does it leave for the next? Check causal and chronological links for an experience-sharing talk, or problem–method–evidence–conclusion links for a technical argument. Do not force every deck into a personal story or rewrite an approved structure solely to fit a template. Keep essential claims and their scope on the slide or in deliberate spoken explanation; speaker notes should not conceal a missing argument.

Useful page roles include cover, orientation, evidence, explanation, comparison, process, decision, boundary, next action, appendix, and closing. These are a catalog, not an exhaustive or exclusive taxonomy.

## 3. Use a modular project

For multi-slide decks, prefer:

```text
slides/
├── compile.js
├── theme.js
├── layout.js
├── slide-01.js
├── slide-02.js
├── imgs/
└── output/
```

Each slide module should export a synchronous function:

```javascript
function createSlide(pres, theme, helpers) {
  const slide = pres.addSlide();
  // Add content.
  return slide;
}

module.exports = { createSlide };
```

Keep shared colors, fonts, margins, typography, citations, page furniture, and reusable components in `theme.js` or `layout.js`. Do not copy theme literals into every slide.

Use semantic tokens that match the deck, for example `background`, `surface`, `text`, `muted`, `primary`, `accent`, `positive`, and `warning`. Add or omit tokens as needed.

Keep one authoritative editable source in the authorized workspace. If later edits are made directly to the PPTX, update the generator or explicitly designate that PPTX as the new editing baseline; do not silently regenerate from stale code. Temporary renders are not the maintainable source.

## 4. Generate assets before assembly

Resolve images, charts, icons, and data before compiling the final deck. Keep source URLs or provenance beside the asset manifest when external evidence is used. Avoid asynchronous network work inside `createSlide()`.

Use the tested image-sizing and endpoint-based line helpers documented in [pptxgenjs.md](../pptxgenjs.md) instead of guessing API options or supplying negative shape extents. Preserve image proportions unless stretching is intentional, and verify that required notes survive generation and reordering.

## 5. Compile predictably

Compile slides in an explicit order and await `writeFile`:

```javascript
const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";

for (const modulePath of orderedSlides) {
  require(modulePath).createSlide(pres, theme, helpers);
}

await pres.writeFile({ fileName: outputPath });
```

Run the script with:

```bash
<skill-dir>/scripts/run_pptxgenjs.sh slides/compile.js
```

## 6. Preview and integrate

When visual direction is unsettled and consequential, preview a small, representative selection of actual-content slides before applying it across the deck. Use [design guidance](design.md) to assess hierarchy, density, color roles, and layout. This is not a mandatory approval gate for routine edits or an already agreed style.

Single-slide previews are useful during difficult layout work. Final judgment must still use the complete deck because navigation, repetition, terminology, and visual rhythm are cross-slide properties.

Use parallel slide authors only when the deck is large and page boundaries are genuinely independent. Keep shared theme, outline, final integration, and visual QA under one owner to prevent style drift.

## 7. Validate

Read [qa.md](qa.md). Text extraction alone cannot detect visual failures.
