# GridForge v2.0 — Canonical Project Handover and Continuation Guide

**Challenge:** GDG Tech Wildcard Entry, Frontend Task 2 — Spreadsheet Web Application. **Checkpoint:** 19 October 2026. **Status:** v2 source built; TypeScript `npm run check` passes; **42/42 automated tests** pass; reproducible engine benchmark passes; offline assembled Chromium browser smoke passes on the exercised interactions. GitHub repository creation, external deployment and a full normal-browser acceptance run must be completed by the user.

## How to resume

Extract the v2 ZIP to any folder (Windows supports ordinary `Downloads`). With Node 20+ installed, double-click `run_windows.bat` and open `http://localhost:4173`; the ZIP includes current compiled `dist/`. To develop:

```powershell
npm install
npm run check
npm test
npm run benchmark
npm run dev
```

No database, backend credentials, environment variables, API keys, or user data files are needed. GitHub publication steps are in `docs/GITHUB_RELEASE.md`.

## Current feature and API inventory

**Original v1 preserved:** 20,800-cell virtual address space, viewport virtualization, 56 custom functions, tokenizer, Pratt parser, dependency graph, selective recalculation, circular-reference errors, syntax-position diagnostics, formula bar, range selection, keyboard navigation, drag-fill, undo/redo, JSON/CSV, What-if, Inspector, Impact.

**New v2 implemented:** independent worksheet tabs (add, rename, duplicate, delete, switch), editable workbook title, sparse per-cell formatting with font weight/italics/underline/colors/alignment, General/Number/Integer/USD/Percent numeric display, style undo/redo, same-app raw formula + style clipboard paste, values-only paste, selected-range ascending/descending sort, Find/Replace with formula protection, Data Quality Lens and bar chart, versioned multi-sheet JSON with legacy v1 migration, CI and GitHub Pages workflow files.

**Exact source responsibilities**

| File | Owner responsibility |
| --- | --- |
| `src/engine/parser.ts` | Lexer, Pratt parser, reference expansion/shifting, formula diagnostics |
| `src/engine/functions.ts` | 56 function registry entries, coercion/error helpers |
| `src/engine/workbook.ts` | Raw/source cells, sparse styles, dependency/reverse dependency maps, cached evaluation, undo/redo, multi-edit transaction, What-if, JSON workbook serialization |
| `src/engine/format.ts` | `CellStyle`, safe style validation, `styledDisplay` numeric presentation |
| `src/engine/document.ts` | `SpreadsheetDocument`, 12-sheet cap, tabs, names, duplication, deletion, v1/v2 document serialization |
| `src/engine/operations.ts` | Range search/replace, stable sort, selection stats, chart-ready data |
| `src/engine/csv.ts` | CSV parsing and quoting/serialization |
| `src/main.ts` | All DOM event wiring, virtual grid, toolbar, side panels, clipboard, file import/export, debounced local persistence |
| `public/index.html`, `public/styles.css` | Static shell, interface and responsive presentation |
| `scripts/benchmark.mjs` | Deterministic 10,000-formula benchmark and expected graph reach checks |
| `scripts/copy-static.mjs`, `scripts/serve.mjs` | Compile static public shell, run lightweight local HTTP server |
| `tests/workbook.test.mjs`, `tests/upgrade.test.mjs` | 26 original + 16 v2 Node engine tests |
| `.github/workflows/{ci,pages}.yml` | Automated static checking/tests/benchmark and optional GitHub Pages deployment |
| `dist/` | Regenerated JavaScript, CSS, HTML for web/Windows launcher and GitHub Pages |

## Exact checkpoint checks

- Node.js `v22.16.0`, TypeScript `5.8.3` CLI.
- `npm run check`: PASS.
- `npm test`: **42 passed, 0 failed**; v2 test suite additionally covers styles/undo, multiple sheets/legacy migration, sort/find/replacement, data analytics.
- `node scripts/benchmark.mjs`: **10,400 populated (10,000 formulas)**, initial load/recalc **95.516 ms**, change A2 **26 affected/recalculated / 0.049 ms**, isolated edit **1 affected/recalculated / 0.007 ms** on this environment. These are local engine timings, not UI rendering or cloud performance.
- Chromium assembled-page smoke (direct HTTP navigation blocked in environment): **231 virtual cell DOM nodes** at 1440×900; D4 initial 4000, updating B4 to 225 made D4 = 7200, undo restored it; Bold plus undo/redo; an added second sheet contained A1=42 without changing Sheet 1; Find `P-103` yielded two matches; custom MIME clipboard translated `=B4*C4` from D4 to H4 as `=F4*G4` and undo cleared H4; Analyze A4:B11 reported sum 951 and average 118.875; sorting and undo worked; downloaded v2 JSON exported both worksheets and formatting; no uncaught JS exceptions. **localStorage could not be tested in the assembled `about:blank` context**. Optional reproduction script: `scripts/browser_smoke_optional.py`; see `docs/TESTING_AND_DEMO.md`.
- Browser screenshot stored at `docs/preview_v2.png`.
- npm registry was unreachable from the packaging runtime; `package-lock.json` could not be reliably generated. GitHub workflows use `npm install`, not `npm ci`. Users may generate and commit lockfile after installing on their own machines, then switch CI to `npm ci`.

## Data formats and migration

- Current browser key: `gridforge:document:v2`; older key: `gridforge:greeting:workbook:v1` (auto migrated if no v2 saved document).
- `SpreadsheetDocument.toJSON()` version 2 holds `format`, `activeId`, and sheet entries `{id,name,workbook}`. Workbook v2 stores `rows`, `cols`, `cells`, and `styles`. Export JSON before changing browsers/ports or clearing local storage.
- Legacy standalone v1 workbook JSON can be imported; it becomes a one-sheet v2 document.
- CSV import/export applies to one sheet, values only; **not** to Excel `.xlsx` packages.
- No schema, database or server; all computed formulas and source data stay in client runtime/browser storage.

## Architecture invariants / do not regress

1. Keep the handwritten formula parser/evaluator. Do **not** use an existing spreadsheet calculation engine.
2. Preserve relative/absolute `$` reference behavior and cycle detection.
3. Keep calculation graphs per sheet, isolated from visual formatting or chart calculations.
4. Preserve 150-operation undo/redo history within a sheet session; formatting must be undoable.
5. `IF` and `IFERROR` must evaluate lazily. Preserve AST diagnostics.
6. Keep browser visible cell DOM count tied to viewport, not total sheet size.
7. Sanitize imported cell styles to strict `#RRGGBB` strings. Use escaped HTML or textContent for external content.
8. `What-if` must clone state; previews may not modify live workbooks.
9. Do not claim Excel parity: sorting and clipboard semantics are documented approximations.
10. Rebuild with `npm run build` whenever changing source; `dist/` is checked into the packaged version for easy demo.

## Known limitations and next milestones

- Cross-sheet formula references (`OtherSheet!A1`) are not implemented. Each sheet recalculates in isolation.
- No `.xlsx` import/export, chart object persistence, filtering, conditional formatting, merged cells, pivot tables, macros or collaborative editing.
- Sorting works on selected rectangle and shifts relative refs in moved formulas; review complex inter-row references.
- Clipboard custom MIME depends on browser permissions. Plain TSV fallback can lose formula/source style metadata outside same-app copying.
- No tab-wide undo history persisted after browser reload. Saving uses localStorage with quotas, not IndexedDB.
- Main-thread recalculation might become slow with dense complex dependencies. Web Workers are future work.
- GitHub Actions and GitHub Pages workflows have not been exercised in a pushed GitHub repository; perform the final push + deployment validation yourself.

**Suggested v3 sequence (not yet implemented):** 1) confirm `npm install` and full browser acceptance on user's machine, 2) push repository and verify CI/Pages, 3) optional XLSX interoperability library strictly for file I/O while retaining custom formulas, 4) cross-sheet formula grammar and dependency graph, 5) richer charts/data validation/filtering, 6) background calculation worker. Do not claim unimplemented items as current functionality.
