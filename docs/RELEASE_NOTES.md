# GridForge v8.0 release notes

- Added **100 distinct Excel-style Data Studio operations** across 10 categories, searchable with preview and one-action undo/redo.
- Added right-click entry, confirmation before replacing populated aggregate result cells, formula-preserving cleanup, relative formula shifting in fill/rearrangement, 12K-cell selection guard and square-range validation.
- Added 107 automated test cases (100 specific action tests and 7 cross-cutting checks), for 447 total unit tests.
- Added 10,000-cell stress test and v8 Data Studio browser regression checks.
- Fixed duplicate-row comparison of identical formula source strings, initial action selection in the modal, and outdated footer version label.
- Preserved 306 formula functions, original v7 XLSX/XLSM and GridForge macro support and optional trusted Windows Excel bridge.
- Documented limitations and outstanding Windows storage/VBA acceptance tests.

---

# Release notes

## v4.0.0 — 80 extra Excel-style functions (136 total)

- Added 24 math/trig, 16 statistics/regression, 12 text, 16 date/time, 8 logic/lookup/conditional and 4 financial functions.
- Implemented lazy `IFS`, `SWITCH` and `XLOOKUP` not-found evaluation.
- Fixed `LOG10(...)` function / cell-reference tokenizer ambiguity without changing the handwritten parser architecture.
- Improved modern Excel `_xlfn.` XLSX import/export handling.
- Added 91 automated tests to the original 65 (156 total) and one new browser check (15 total).
- Preserved the existing Python-only launcher, virtualized grid, dependency graph, sorting, formatting, filtering and multi-sheet workbook.
- Full reference: `docs/FUNCTION_REFERENCE.md`; category guide: `docs/V4_FUNCTION_EXPANSION.md`.


## v3.0.2 — Final GitHub submission and Python-only setup

- Adds Python 3 standard-library local web server so the compiled app can run with **no installed npm or pip modules**.
- Adds Windows one-click launcher and macOS/Linux launch script.
- Adds Python static-site smoke test and integrates it into GitHub CI.
- Includes verified GDG requirement mapping, fresh setup walkthrough, reviewer demo, release evidence, corrected GitHub instructions and an updated handover.
- Keeps TypeScript spreadsheet engine, browser features and bundled compiled `dist/` unchanged relative to v3.0.1.
- Historical v3.0.1 Windows evidence: 65/65 Node tests and 14/14 browser checks passed.
- Native XLSX interoperability is a supported subset; cross-sheet formula evaluation and full Excel parity are still not supported.

## v3.0.1 — Regression hardening

Expanded tests to 65 automated Node cases, 14 Chromium browser checks, full regression orchestrator, benchmark, CI integration and regression docs. No runtime feature changes.

## v3.0 — Excel file support and spreadsheet productivity

Introduced custom Office Open XML XLSX import/export, basic table filtering, freeze row/column, numeric conditional formatting, AutoSum and import diagnostics on top of v2.1 features.

## v2.1 — Context menu and table sorting

Introduced Excel-inspired custom right-click, sort-by-column with header handling, formatting-preserving sorting, cut/paste operations and browser tests.

## v2 — Everyday spreadsheet tools

Multiple tabs, cell styles, number formats, find/replace, sorting, analysis and charts.

## v1 — Custom formula engine

Virtualized grid, 56 custom functions, Pratt parser, incremental dependency tracking, formula diagnostics and Undo/Redo.
