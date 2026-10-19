# GridForge v2.0 — Release Notes

Date: 19 October 2026. Scope: GDG Tech Wildcard Entry, Frontend Task 2.

## Added

- Multi-worksheet document container; add, switch, rename, duplicate, delete up to 12 worksheets.
- Per-cell formatting: bold/italic/underline, font and fill hex colors, alignment, display-oriented number formats.
- Undoable formatting changes and combined raw-content plus style paste transaction.
- Formula-aware same-application clipboard with relative/absolute address shift and plain TSV interoperability.
- Paste Values from the most recent internal clipboard copy.
- Selected-range ascending/descending sorting, numeric key handling, shifted relative formulas.
- Find All, clickable results, protected Replace All (raw formulas opt-in).
- Data Quality Lens with selection statistics and inline bar chart.
- v2 multi-sheet JSON format; legacy v1 workbook import and localStorage migration.
- GitHub Actions CI, GitHub Pages deployment workflow, repository deployment guide, optional browser smoke script.

## Preserved

- Entire v1 handwritten formula parser, 56 functions, dependency tracking, caching, cycle and parsing errors, What-if, Impact and Inspector.
- 400 × 52 default grid, virtualization, formula bar, editing/selection, drag fill, per-sheet undo/redo, CSV support.

## Verification at release

- `npm run check` succeeded, TypeScript 5.8.3.
- `npm test`: **42 passing, zero failing**.
- `npm run benchmark`: verified 10,000 formulas, 10,400 populated cells, 26 affected cells on linked edit, one affected on isolated edit.
- Compiled-browser smoke: no page exceptions in the tested paths, plus edit/undo, style/undo, two-sheet isolation, find, custom clipboard, sorting, Analyze and exported two-sheet JSON.

## Not included

No actual Excel `.xlsx` round-trip, cross-sheet formulas, pivot tables, macros, named ranges, true Excel date formats, complex merging, native persistent charts or server-based multi-user collaboration. GitHub Pages workflow is included but no deployed public link has been verified.
