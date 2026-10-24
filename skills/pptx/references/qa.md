# PPTX Quality Assurance

Assume the first render contains problems. Validate content and appearance separately.

## Content QA

```bash
python -m markitdown output.pptx
python -m markitdown output.pptx | grep -iE "\bx{3,}\b|lorem|ipsum|\bTODO\b|\[insert|this.*(page|slide).*layout"
```

Check completeness, ordering, spelling, figures, units, citations, terminology, speaker notes when required, and leftover template content.

## Render safely

```bash
<skill-dir>/scripts/render_slides.sh output.pptx [render-root]
```

The command creates a new render directory for every run and prints the absolute paths of all slide images. Use those exact paths for visual inspection.

## Visual QA checklist

- Text, images, charts, lines, footers, and citations do not overlap.
- Nothing is clipped at shape boundaries or slide edges.
- Long titles and translated text wrap intentionally.
- Margins, columns, card dimensions, and repeated elements align.
- Contrast remains readable on the rendered background.
- Chinese and Latin text use the intended fonts or acceptable verified substitutes.
- Images are sharp enough and are not stretched.
- Charts include units, labels, legends, sources, and readable values where needed.
- Page furniture does not compete with the content.
- The deck has a coherent rhythm without accidental layout repetition.
- Live, replayed, simulated, estimated, and measured evidence is labeled truthfully.

For an important or complex deck, use a fresh independent visual reviewer when available. A subagent is optional for small decks; the independent inspection and evidence are what matter.

## Verification loop

1. Generate or edit the deck.
2. Extract and inspect its content.
3. Render every slide and list observed issues.
4. If there are no observed issues and required checks pass, deliver. Otherwise fix issues within the authorized scope.
5. After edits, regenerate the PDF and slide images in a fresh render directory.
6. Recheck affected slides and scan for cross-slide regressions related to the changes.

Complete when required content, visual and structural checks pass and observed issues within scope are resolved. Zero observed issues is a valid outcome. Re-render after actual edits; do not invent an issue or repeat an unchanged check to satisfy an iteration count. For analysis-only requests, report findings without editing.

## Structural checks

Template editing should use the bundled `pack.py --original` workflow so inherited source-package issues are distinguished from newly introduced errors. For generated decks, also verify that LibreOffice can convert the final PPTX to PDF.

If structural validation and successful application rendering disagree, report the exact validator error rather than silently ignoring it or claiming the file is fully valid.
