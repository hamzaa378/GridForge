# CURRENT HANDOVER: GridForge v10.0.0

v10 adds `src/engine/functions_v10.ts` (100 new callable formula names), persistent light/dark + Excel/Guided settings in `src/main.ts` and `public/styles.css`, pre-paint preference initialization in `public/index.html`, `tests/functions_v10.test.mjs`, v10 browser checks, and a 100-function reference workbook containing 30 verified evaluated formulas. **Current total: 506 registered names** and 100 v8 Data Studio actions. Run `npm run check`, `npm test`, and `npm run regression` with Playwright. Python-only compiled launch: `python run_gridforge.py`. Functions have explicit compatibility limits, especially lambda, dynamic arrays, web data, and Excel VBA. See `V10_RELEASE_EVIDENCE.md`, `V10_THEME_AND_MIGRATION.md`, `V10_FUNCTIONS_AND_LIMITATIONS.md`, and `V10_EXCEL_FUNCTION_GAP.md`. Do not remove compiled `dist/` from GitHub repository or the Python-only launch will fail.

**Next developer:** maintain the version number, modify source and rebuild `dist/`, add meaningful function and browser tests for every feature, verify preference reload on Windows, re-run the XLSM/VBA security checks, and keep this handover and README current. Avoid claiming desktop VBA execution until actually tested on Windows with Excel.

---

# CURRENT HANDOVER: GridForge v9.0.0

Latest source: `src/engine/functions_v9.ts` (100 new functions, bringing total to 406), `src/main.ts` (optional familiar Excel ribbon, command search and Start here), `src/engine/categories.ts` (Database/Compatibility categorization), `tests/functions_v9.test.mjs` (individual behavior checks for all new functions), `tests/v9_examples.mjs` (shareable test/gallery formulas), `scripts/build_v9_gallery.mjs`, and `scripts/build_excel_gap_report.mjs`. Frozen source catalog: `docs/excel_comparison_catalog.json`; generated gap: `docs/V9_EXCEL_FUNCTION_GAP.md`. Run `npm run regression`; standard-library-only Python launch remains `python run_gridforge.py`. The XLSM updater and optional Windows desktop Excel bridge are preserved. See `V9_RELEASE_EVIDENCE.md`, `V9_CODE_AUDIT.md`, `V9_UI_QUICK_START.md` and README for limitations.

---

# Current checkpoint — GridForge v8.0 (100 Data Studio actions)

**This is the primary continuation guide.** The most recent release is v8.0 and extends the v7 XLSM/VBA-preserving project; older sections below are retained as technical history.

## Current state

- Custom TypeScript spreadsheet engine with 306 registered functions, 20,800-cell virtualized default sheet, 12-sheet maximum, formula dependency caching, XLSX and supported XLSM import/export, and optional Windows Excel bridge (VBA execution remains unverified without Windows Excel).
- **New v8 Data Studio:** `src/engine/tools_v8.ts` implements 100 unique, categorized pure preview operations; `src/main.ts` renders the searchable modal and applies changes using workbook transactions with Undo/Redo; `public/styles.css` styles the modal.
- 10 categories of 10 commands, with 100 individual behavior tests, 447 total unit tests, 26 browser checks, 10,000-formula benchmark and 10,000-cell Data Studio stress test. Full pipeline adds one new benchmark stage (7 stages).
- Data Studio source is in `src/engine/tools_v8.ts`; `tests/data_studio_v8.test.mjs` gives expected outputs for every registered action; `scripts/benchmark_tools_v8.mjs` checks 10K changes; `scripts/build_v8_docs.mjs` regenerates the complete function catalog and `examples/v8_data_studio_demo.json`.
- `dist/` contains compiled JS and assets, including `dist/engine/tools_v8.js`; this is essential to keep in the GitHub repository so Python-only launch works.
- `docs/V8_100_ACTION_REFERENCE.md` is the feature catalogue; `docs/V8_CODE_AUDIT.md` contains confirmed defects fixed and open limits; `docs/V8_ACCEPTANCE_CHECKLIST.md` is the human test guide; `docs/V8_RELEASE_EVIDENCE.md` has verification results.

## Build and test commands

```powershell
python run_gridforge.py
npm install
npm run build
npm run regression:core
npm run regression
python scripts/python_smoke.py
```

Full browser regression requires Python Playwright and Chromium, installed using `python -m pip install playwright` and `python -m playwright install chromium`. For trusted VBA only, install desktop Excel on Windows, `pywin32`, and launch `scripts/excel_vba_bridge.py` explicitly.

## Known limits to preserve and explain

1. The 100 new items are **data editing operations**, not 100 Excel formula functions; function count remains 306. Do not claim complete Excel parity.
2. The 12,000-cell Data Studio cap and square-only transpose/rotation are intentional safeguards; text/numeric actions skip formulas.
3. XLSM can preserve VBA binaries for supported cell-content changes but not every formatting/structure/shared-formula modification. Original signed package signatures might not remain valid after edits.
4. Browser macro execution is impossible; real VBA requires the optional trusted Windows Excel desktop bridge. No Windows Excel test has run in this environment.
5. Browser localStorage reload was skipped in hosted Chromium due to navigation restrictions. Check this in a user's normal Windows browser.
6. The original v7 regression tests must remain passing when further features are added. Do not change XLSM security checks or parser semantics without targeted tests.

## Follow-up work

After user-side acceptance, validate with genuine Excel-created XLSX/XLSM files, run Windows VBA acceptance with a trusted test workbook, inspect performance on 12K edited cells under typical hardware, and consider full data validation, row/column resizing, print layout, charts, and cross-sheet formulas as distinct future features. Those are **not** part of v8's 100 actions.

---

# v7 continuation checkpoint — 2026-10-21

Current release: 7.0.0 (built from the tested v6.0.0 baseline). The front-end engine retains 306 functions, 340 unit tests, Python-only compiled-site launch. New modules: `src/engine/xlsm.ts` patches supported raw cell contents/formulas into original OOXML package while preserving binary VBA and unrelated parts; `src/engine/zip.ts` now supports binary-preserving ZIP reads/writes; `scripts/excel_vba_bridge.py` runs explicitly named macros using installed Windows Microsoft Excel; `scripts/test_xlsm_browser.py` and `scripts/test_vba_bridge_protocol.py` provide isolated regression/security checks. The Automate dialog now includes an opt-in bridge token and named VBA macro control; `Save XLSM` exports edited macro-enabled files.

**Unverified live dependencies:** The authoring environment is Linux, so Excel COM/pywin32 live execution and full Excel opening of modified XLSM require the Windows test described in `docs/V7_NATIVE_XLSM_AND_VBA.md`. Do not claim it has been executed here. Excel signatures, shared/array formulas, style edits and structural sheet edits have documented restrictions. Macro bridge honours Trust Center and disables Excel events.

**Next steps:** Run on Windows with licensed Excel, test a trusted workbook with `Module1.GridForgeSmokeTest`, verify new XLSM opens and original VBA project remains valid, and run the full Playwright regression locally. Only after acceptance should this release be deployed/published as production.

---

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


## Test-suite checkpoint: v3.0.1

The current packaged repository adds regression coverage only, with no runtime feature changes. Use `npm run regression` for 65 Node automated tests, the 10,000-formula benchmark and 14 browser checks. See `docs/REGRESSION_TESTS.md` and `docs/REGRESSION_HANDOVER.md` for dependencies, commands, diagnostic output, limitations and next steps.


## Final packaging checkpoint: v3.0.2 (Python-only release)

The application/engine source and prebuilt browser output `dist/` are unchanged from the v3.0.1 regression-tested version. The new release adds `run_gridforge.py`, `START_HERE_WINDOWS.bat`, `START_HERE_MAC_LINUX.sh`, `scripts/python_smoke.py`, a setup-from-zero guide, task compliance mapping, reviewer walkthrough, and corrected final README/GitHub instructions. **A person with only Python 3 can run the precompiled app without installing any modules**, while Node is required for TypeScript rebuilds and Node engine tests. Python Playwright and Chromium are additionally required for automated browser regressions. Start with `docs/SETUP_FROM_ZERO.md` and `docs/TECHNICAL_COMPLIANCE.md`.

The supplied Windows test evidence on v3.0.1 is 65/65 unit tests, 14/14 browser checks, and 4/4 regression stages. The v3.0.2 Python static server was smoke tested separately; the full 65-unit/14-browser combined regression also passed in a hosted Chromium environment, with the browser-storage reload skipped because direct localhost navigation was restricted. The user's real Windows v3.0.1 HTTP regression already verified browser persistence, but v3.0.2 should still be checked after extraction on Windows if an environment-specific release certification is desired. The included `dist/` is the static runtime served by Python, not a Python spreadsheet implementation. All existing v3 Excel compatibility limitations remain.


## GridForge v4.0.0 — continuation checkpoint (20 October 2026)

- Baseline: the original tested v3.0.2 frontend and 56-function custom engine.
- Added file: `src/engine/functions_extra.ts` containing 80 additional functions; registration in `src/engine/functions.ts` yields 136 functions.
- New evaluation paths: `IFS`, `SWITCH` and `XLOOKUP` optional fallback in `workbook.ts`; `LOG10` tokenization correction in `parser.ts`.
- OOXML compatibility: normalizes `_xlfn.` extended Excel names in `xlsx.ts`; emits prefixes for supported newer functions.
- Reproducible tests: `tests/functions_v4.test.mjs`, updated `scripts/regression_browser.py`, and original engine, sort, data, XLSX and browser tests.
- Documentation: `docs/V4_FUNCTION_EXPANSION.md`, `docs/FUNCTION_REFERENCE.md` (generated), `docs/V4_RELEASE_EVIDENCE.md`.
- Build: Node 20+ / TypeScript 5.8, `npm install`, `npm run build`, `npm run regression` (browser needs Python Playwright/Chromium). Pure-Python launch: `python run_gridforge.py` serves bundled `dist/`.
- Known limits: dynamic arrays and cross-sheet formulas remain unsupported; XLSX import/export supports a subset only; date/time serials use UTC; volatile functions update on recalculation.
- Changes to function signatures should include additions to `tests/functions_v4.test.mjs` and regeneration with `npm run functions:docs`.
- Do not claim cross-sheet formula support or full Microsoft Excel feature parity. Browser storage reload remains a separate Windows acceptance task for v4.

---

## GridForge v5.0.0 — current release and continuation checkpoint

**Baseline:** v4.0.0; no changes to the handwritten parser architecture or the dependency graph semantics.

**New components:** `functions_v5.ts` adds 80 independently registered functions (216 total); `categories.ts` groups the full registry; `macros.ts` validates and replays safe structured cell/style actions. `main.ts` adds category filter, Automate controls and in-memory byte-identical download of the imported `.xlsm` original. `xlsx.ts` adds dotted Excel name aliases for selected formulas. Embedded VBA remains non-executable. Edited exports are `.xlsx` without VBA, and original `.xlsm` never receives edits.

**Test evidence:** 241 Node unit tests passed; 21 browser checks passed in hosted embedded mode (storage reload skipped). The 10,000-formula benchmark and Python-only smoke are part of the final regression. Actual user Windows test is still pending for v5; previous user checks applied to v3.

**Build commands:** `npm install && npm run build`; `npm run regression:core`; `npm run regression`; `npm run gallery:build`. Launch compiled app with `python run_gridforge.py` (no pip modules) or `START_HERE_WINDOWS.bat`.

**Files:** `examples/v5_80_function_gallery.xlsx`, corresponding JSON, generated `docs/FUNCTION_REFERENCE.md`, `docs/V5_FUNCTIONS_AND_MACROS.md`, `docs/SECURITY.md`, updated README/CI. Keep `dist/` in release ZIP for Python-only use.

**Known limitations:** No VBA interpreter or XLSM edit merge; `.xls` binary unsupported; cross-sheet formulas, full Excel formats, dynamic arrays and Excel Object Model unavailable. Macro recording covers edited cells and formatting only. Financial/date/regex functions are documented, bounded subsets and require independent validation for real decision-making.

**Next work if continuing:** Verify full localStorage reload on user's Windows PC; run import/export against genuinely macro-enabled Office-created `.xlsm` documents without enabling VBA; extend formulas using independent Excel outputs; add deeper workbook interoperability only when a safe, testable preservation mechanism exists.

## GridForge v6.0.0 — current checkpoint

**Baseline:** cloned from verified v5.0.0. Adds `src/engine/functions_v6.ts` (90 new functions; total **306**), parser support for dotted names, and runtime dependency tracking for `INDIRECT` / scalar `OFFSET`. The function browser now highlights the 90 recently added v6 functions.

**Automation:** `src/engine/macros.ts` now validates both v1 and v2 JSON. V2 can target multiple existing worksheets and handles `cells`, `styles`, `clear`, `fillDown`, `copy`, `sort` and `replace`. Dry-run preflight simulates all steps on a document copy, returns a diff and prevents partially committed changes when any step is invalid. Recorded cell/style edits now capture the active worksheet; Custom Sort is recorded. UI provides preview and remembers the selected macro after reopening the dialog. VBA is still never executed and changes are never merged back into `.xlsm`.

**Files:** `tests/functions_v6.test.mjs` tests each of the 90 functions; `tests/macros_v6.test.mjs` checks unsafe input, multi-sheet replay, dry-run, old v1 compatibility, and rollback. `scripts/build_v6_gallery.mjs` generates `examples/v6_90_function_gallery.xlsx` and JSON. `docs/V6_NEW_90_FUNCTIONS.md`, `docs/V6_AUTOMATION_AND_XLSM.md`, `docs/V6_RELEASE_EVIDENCE.md` and `docs/V6_REVIEWER_CHECKLIST.md` explain usage and evidence. Rebuild function reference via `npm run functions:docs`.

**Commands:** Python-only `python run_gridforge.py`; developer setup `npm install`, `npm run dev`; tests `npm run check`, `npm test`, `npm run benchmark`, `npm run regression`. Run Windows browser tests after installing Python Playwright+Chromium. The compiled `dist/` must be included for Python-only use. **Do not assume macros in imported Excel workbooks run.**

**Measured locally:** 340 Node tests passing; 24 browser checks passing (storage-reload skipped on this hosted browser), 10,000-formula benchmark passing; Python-only smoke check to be repeated after ZIP assembly. v6 improves feature breadth but does not claim full Excel compatibility.

## 2026-10-23 — v10.6 Macro Studio checkpoint

Baseline: v10.5 Windows desktop build kit (506 local functions, 10 external placeholders, Excel-style ribbon, native macro v3, optional VBA bridge). Added v4 typed parameters, text cleanup, conditional styling, duplicate-range compaction; source `src/engine/macros.ts`. Added JSON parameter input, new templates and delete option in `src/main.ts`. Added `scripts/run_macro_offline.mjs` for Node-based CLI, `examples/macros/*.json`, separate Windows Start menu Macro Studio shortcut (`--macro-workbench`) in `desktop/gridforge_desktop.py` and Inno Setup installer. Tests: 847 TypeScript tests passed, Python desktop tests and benchmark separately; browser automation is blocked by the hosted browser policy and needs Windows smoke testing. Build kit only, not a compiled EXE. See `docs/MACRO_STUDIO_V4.md` and `docs/DESKTOP_BUILD_AND_RELEASE.md`. No native VBA interpreter is present.
