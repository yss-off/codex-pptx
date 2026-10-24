---
name: pptx
description: "Create, read, analyze, edit, combine, or validate PowerPoint presentations. Use whenever a task involves a .pptx file, slide deck, presentation template, speaker notes, slide layouts, or presentation QA."
license: MIT
---

# PPTX Skill

## Choose the workflow

| Task | Required guidance |
|------|-------------------|
| Read or analyze content | Use `python -m markitdown`; render when visual structure matters |
| Edit or create from a template | Read [editing.md](editing.md) |
| Create from scratch | Read [references/generation.md](references/generation.md) and [pptxgenjs.md](pptxgenjs.md) |
| Choose visual direction or layouts | Read [references/design.md](references/design.md) |
| Prepare a first-touch technical or customer alignment deck | Read [references/technical-alignment.md](references/technical-alignment.md) |
| Validate any final deck | Read [references/qa.md](references/qa.md) |

Resolve commands relative to this skill directory. Use absolute script paths when the working directory is elsewhere.

## Read and inspect

```bash
python -m markitdown presentation.pptx
python <skill-dir>/scripts/thumbnail.py presentation.pptx
python <skill-dir>/scripts/office/unpack.py presentation.pptx unpacked/
```

Text extraction verifies content, not layout. Render the deck whenever placement, typography, charts, images, or template fidelity matter.

## Core workflow

1. Establish the audience, presentation purpose, decision or belief sought, delivery format, and source-of-truth content.
2. Choose template editing or from-scratch generation. Preserve the original input file.
3. Plan one clear job for each slide. Treat page-type and layout catalogs as options, not mandatory taxonomy.
4. Build with shared theme and layout helpers. Keep content, evidence, citations, and speaker notes traceable to their sources.
5. Run content QA and visual QA. Fix observed issues within scope and re-render affected slides after changes. If all required checks pass with no issues, deliver without making an artificial edit.
6. Deliver the `.pptx` plus any specifically requested supporting files; do not leave previews or temporary artifacts beside the final file.

## Defaults, not universal rules

- Prefer a visual anchor on content slides, but allow typography-led or intentionally sparse slides when that better serves the message.
- Prefer varied layouts, but preserve repetition when it improves comparison or navigation.
- Use readable size contrast and generous spacing, adjusted for the actual template, room, and display conditions.
- Page numbers, badges, gradients, animations, and section dividers are optional unless the user or template requires them.
- Use a coherent palette and semantic theme tokens; do not force a fixed number or naming scheme when the deck needs more.
- Keep body text left-aligned by default. Center text only when the layout purpose supports it.

## Environment and reliability

- Run PptxGenJS files through `scripts/run_pptxgenjs.sh`; global npm packages are not automatically resolvable by Node on every host.
- Detect installed fonts before choosing them. For Chinese decks, prefer a CJK font available in the rendering environment and verify the rendered result. Do not assume Microsoft YaHei exists.
- Use `scripts/render_slides.sh` for visual QA. It creates a fresh render directory and avoids deleting wildcard files from the working directory.
- For template edits, use the bundled unpack, add-slide, clean, pack, and validation scripts instead of rebuilding fragile package relationships ad hoc.

## Completion checks

- Extracted text is complete, correctly ordered, and free of placeholders.
- Rendered slides have no clipping, overlap, unintended substitution, stale images, or unreadable contrast.
- Claims, figures, units, citations, terminology, and live-versus-replay labels are consistent.
- The final PPTX opens through LibreOffice conversion; template edits also pass the bundled pack/validation workflow or report any known validator disagreement explicitly.
- A fresh render after the final edit shows no new issues.

## Dependencies

- `markitdown[pptx]`
- Pillow
- PptxGenJS
- LibreOffice
- Poppler (`pdftoppm`)
- Optional for icons: `react-icons`, `react`, `react-dom`, `sharp`
