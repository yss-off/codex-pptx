# PPTX Quality Assurance

Look for problems in the first render; do not equate successful rendering with quality. Keep the following judgments separate:

| Check | Evidence it needs |
|---|---|
| Content and narrative | Approved brief, displayed slide order, page text and notes; coherent title sequence, supported claims, and useful transitions |
| Visual quality | Full-page renders plus a deck overview; hierarchy, readability, color, alignment, and repetition |
| Structure and editability | Source-package validation, relationships, native objects, and notes matched to their displayed slides |
| Live delivery | Actual rehearsal or audience/projection feedback; render checks alone cannot establish timing or comprehension |

Report definite defects, likely misunderstanding, aesthetic preferences, and unverified conditions distinctly. A positive judgment in one category is not a pass in the others.

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

When an independent review is used, provide the current render, notes, and confirmed constraints. Reconcile suggestions with those constraints and verify concrete findings locally; do not silently apply every recommendation or make external/Pro review a default requirement.

## Verification loop

1. Generate or edit the deck.
2. Extract and inspect its content.
3. Render every slide and list observed issues.
4. If there are no observed issues and required checks pass, deliver. Otherwise fix issues within the authorized scope.
5. After edits, regenerate the PDF and slide images in a fresh render directory.
6. Recheck affected slides and scan for cross-slide regressions related to the changes.

Complete when required content, visual and structural checks pass and observed issues within scope are resolved. Zero observed issues is a valid outcome. Re-render after actual edits; do not invent an issue or repeat an unchanged check to satisfy an iteration count. For analysis-only requests, report findings without editing.

## Structural checks

Newly generated decks need full validation, without an original-file exemption:

```bash
python <skill-dir>/scripts/office/validate.py output.pptx
```

Template editing should also use `pack.py --original` to distinguish inherited issues from introduced errors. `NO_NEW_ERRORS` means no new distinct XSD diagnostic messages per part; it does not compare the number or location of repeated error instances. This baseline result is not full validity: report remaining inherited errors and skipped checks. Resolve in-scope generation defects rather than repeatedly accepting them as a baseline. A required change outside the authorized scope should be reported, not silently repaired. Verify LibreOffice conversion independently of the schema result.

If structural validation and successful application rendering disagree, report the exact validator error rather than silently ignoring it or claiming the file is fully valid.

For local revisions, verify the promised scope: a palette edit preserves text, geometry, notes and order; a layout edit preserves intended text and notes; reordering preserves each slide's notes and updates any visible page numbers. Resolve the presentation's relationship-based display order, not numeric `slideN.xml` filenames. Compare only the invariants appropriate to the requested change, then inspect affected renders and the overall deck.

For revisions retaining source part identities, use the scoped comparison helper:

```bash
python <skill-dir>/scripts/compare_decks.py before.pptx after.pptx --mode palette --slides 4,7,11
```

`--slides` uses 1-based display order; omit it to allow changes throughout the deck. `layout` allows geometry and styling edits inside existing slide/layout/master XML, while preserving text, content references, playback settings, notes and relationships. Layout switches or dependency rewrites need separate review. Global style changes are conservatively rejected for a partial slide selection. `palette` compares existing color slots; adding an explicit override to inherited color may be conservatively rejected. `reorder` checks a pure order change without `--slides`; deliberate visible-page renumbering needs separate review. These structural checks do not prove visual equivalence or compare independently rebuilt decks by meaning.
