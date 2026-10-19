# MAIN HANDOVER — GridForge v3.0 (GDG Frontend Task 2)

**Resume here first.** This guide is the authoritative current state, not a historical v2.1 plan.

## Current state

- **v3 source ready:** `src/`, `public/`, `.github/workflows/`, compiled `dist/`, `examples/sample_sales.xlsx`.
- **Task:** custom TypeScript formula parser/evaluator, minimum 30 functions (implemented 56), dependency-aware incremental recalculation, 10,000+ virtualized cells, circular detection, diagnostic display, undo/redo, fill handle, formula bar, keyboard navigation.
- **v2.1 kept intact:** 12 worksheet tabs, formats, sorting, custom context menu, search/replace, copy/paste, dependency and inspector panels, analytics, scenario mode.
- **New v3:** Office Open XML XLSX import/export, shared-string/shared-formula handling, filter selected table rows by substring, freeze top row/first column, live numeric conditional formatting, AutoSum and import report.
- **No formula engine dependencies:** `parser.ts`/`functions.ts`/`workbook.ts` remain handwritten. `xlsx.ts` and `zip.ts` are only file interoperability and archive utilities. No new npm runtime dependency.
- **Persistence:** localStorage key stays `gridforge:document:v2` to pick up local prior v2.1 workbooks on **the same browser origin**. Document JSON is now version 3; version 2 imports automatically. Changing localhost port changes origin and therefore storage. Explicitly export JSON when migrating.

## Important modules

| File | Responsibility |
| --- | --- |
| `src/main.ts` | Shell UI, events, virtualized grid, render, frozen panes, filtered row mapping, conditional colors, import/export dialogs |
| `src/engine/xlsx.ts` | OOXML mapping between GridForge sheets and Excel worksheet/style/relationship XML, import diagnostics |
| `src/engine/zip.ts` | CRC32-verified stored ZIP writer, central-directory parsing, deflate-raw reader |
| `src/engine/sheettools.ts` | Validated persisted filter, freeze and conditional-rule metadata; matches and comparison |
| `src/engine/document.ts` | Named multi-sheet workbook with v2 compatibility and v3 metadata |
| `src/engine/workbook.ts` | Sparse cells, parser ASTs, dependency/reverse dependency graph, recalculation, transactions |
| `src/engine/parser.ts` | Handwritten tokenizer, Pratt parser, diagnostics, formula reference shifting |
| `src/engine/functions.ts` | 56 implemented custom spreadsheet functions |
| `src/engine/operations.ts` | Sort by any selected key with header handling, search/replace, analytics |
| `src/engine/clipboard.ts` | Copy/paste snapshot, relative-formula-aware edits and undoable moves |
| `public/styles.css` | Dark responsive application UI and dialogs |
| `scripts/browser_v3_smoke_optional.py` | Optional full Chromium UI smoke, XLSX DEFLATE roundtrip, screenshot |
| `scripts/benchmark.mjs` | 10,000 formula experiment |
| `scripts/generate_sample.mjs` | Rebuild `examples/sample_sales.xlsx` via GridForge XLSX writer |

## Run/check/test

```bash
npm install
npm run dev                 # prints server URL, defaults to 4174 with fallback
npm run check
npm test                    # 57 passing automated tests at handover
npm run benchmark           # 10,000 formulas, isolate row -> 26 affected
npm run build
node scripts/generate_sample.mjs
```

Optional local Chromium browser regression: set `GRIDFORGE_CHROMIUM=/usr/bin/chromium` on Linux, install Python Playwright, run `python scripts/browser_v3_smoke_optional.py`. The browser script uses Playwright's Chromium binary override. On Windows set the environment variable to your Chromium executable (or remove it if Playwright browsers are installed). This script is not part of required CI.

**Measured here:** TypeScript check PASS, 57/57 Node tests PASS, Chromium browser regression PASS with no uncaught page errors, ZIP validation PASS, local performance benchmark PASS (10,000 formulas, changed row affected 26). Results are local evidence only, not production SLA or external Excel validation.

## Architectural safety rules

- Do not add any library that evaluates formulas: that would violate the challenge rule.
- Keep v2 JSON deserialization compatible; use **v3** for new serialized document format, preserve worksheet metadata across duplicates.
- Do not silently discard unsupported spreadsheet imports or promise `.xlsx` lossless interoperability.
- Imported XLSX is bounded to <=25MB archive, <=12 sheets, <=1000×104 grid per sheet. For larger workbooks, display an error; do not truncate.
- Never execute VBA or external links; `<f>` content is data until parsed by the local GridForge engine.
- Filter is a **view**, not a data rewrite; clearing filter restores unchanged worksheet cells.
- Conditional rules apply to numeric evaluated values; styled background takes priority only when no conditional rule matches.
- XLSX export requests Excel recalc; JSON remains the authoritative GridForge feature backup.

## Known limitations / next milestones

1. Excel compatibility is a subset: no pivot tables, merges, Excel charts, macro execution, date/time fidelity, true Excel data validation, advanced filters or complex conditional rules.
2. Parser and dependency graph are **sheet-local**; quoted/unquoted `Sheet2!A1` references are not yet evaluated. Full cross-sheet dependency graph and cross-sheet cycles would be a worthwhile next research milestone.
3. Freeze is limited to top row and first column (not arbitrary split); no resize of rows/columns.
4. Text-contains filter is one selected column per sheet at a time; OR/AND multi-key filtering is future work.
5. XLSX writer uses stored ZIP (larger output). Compressed XLSX input uses browser `DecompressionStream`; old browser versions might not support it.
6. No backend/sync; localStorage is origin-specific and quota-limited. CSV and XLSX are interoperable subsets; for full GridForge reload prefer JSON.
7. No external spreadsheet app validation was performed in this checkpoint; Windows Excel/LibreOffice user test should be recorded separately before making compatibility claims.

**Next recommended engineering steps after user feedback:** perform Windows acceptance checklist; review import of a real Excel file with shared formulas and dates; integrate cross-sheet references with deterministic cycles and dependency propagation; consider multi-column filter/dropdown validation, frozen arbitrary panes, accessible headers and wider grid navigation. Keep benchmark/test evidence current.
