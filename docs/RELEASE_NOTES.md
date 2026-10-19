# GridForge v3.0 (2026-10-19)

- Added native multi-sheet XLSX import/export using self-authored OOXML ZIP adapter. Browser import handles stored/DEFLATE ZIP, shared strings, shared formulas, supported styling and sheet metadata. Excel formula evaluation still uses GridForge’s own parser.
- Added one-column non-destructive table filtering with row virtualization, freeze top row/first column, live numeric conditional formatting, AutoSum and custom right-click extensions.
- Extended v3 JSON with sheet tools while retaining v2 JSON import. Added an XLSX import report and a reproducible example workbook.
- Baseline: TypeScript PASS, 57/57 automated tests PASS; browser regression and compressed XLSX round-trip PASS, 10,000 formula benchmark PASS.
- Limitations: Excel subset, not complete compatibility. Sheet-local formulas only; advanced Excel features unsupported.

---

## Historical notes from v2.1 and earlier

# v2.1 — Context menu & custom sorting

- Excel-inspired accessible right-click menu on virtualized cells.
- Right-click respects rectangular selection; keyboard Shift+F10 and Up/Down/Escape supported.
- Editing, clipboard, styling, searching, analytics and formula inspection from the menu.
- Header-aware Custom Sort dialog with a selectable key column inside the range.
- Sorting values and cell formatting in a single undoable transaction.
- Internal Cut is deferred until successful Paste; overlapping moves and cross-sheet safety tested.
- Default server port 4174 with automatic fallback to 4184 when PORT is unspecified.
- Added separate sort/clipboard tests and browser interactions to the optional smoke script.
- Formula engine, workbook JSON v2 serialization and all 56 functions unchanged.

## Limitations

This is not full Excel parity. `.xlsx` is not supported. Selected ranges are not auto-expanded on sort. Complex formula rewriting during sort or cut/move is simplified. Cut across sheets is implemented as copy to avoid unintentional deletion. Context Paste uses the internal clipboard snapshot; Ctrl+V accepts external TSV. Browser clipboard access may require user permissions.

---

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
