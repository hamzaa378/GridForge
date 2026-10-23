# GridForge v3 testing update

**Current automated baseline is 57 Node tests, plus the optional `scripts/browser_v3_smoke_optional.py`. v3 instructions are in `docs/V3_ACCEPTANCE_CHECKLIST.md`, `docs/V3_XLSX_IO.md` and `docs/PROJECT_HANDOVER.md`. The older v2.1 evidence below is historical.**

## v2.1 context menu and custom sort

See `V2_1_ACCEPTANCE_CHECKLIST.md` for reproducible manual testing and `V2_1_CONTEXT_MENU_AND_SORT.md` for sorting/cut semantics. Additional Node tests in `tests/context-sort.test.mjs` and additional Chromium interactions in `scripts/browser_smoke_optional.py`.

# Testing, Acceptance Checks, and Demonstration

## 1. Environment setup

```powershell
npm install
npm run check
npm test
npm run dev
```

Open `http://localhost:4174`. Note the Node.js version (`node --version`) and browser version when reporting results. The app itself has no cloud dependencies.

## 2. Automated unit tests

`tests/workbook.test.mjs` and `tests/upgrade.test.mjs` run using Node's built-in test runner (no Jest dependency). The test suite covers these categories:

1. Address capacity >= 10,000, 136 built-in functions and lexer coverage.
2. Operator precedence, right-associated power, unary operators, concatenation and comparisons.
3. SUM, AVERAGE, COUNT, COUNTA, COUNTBLANK, MEDIAN, PRODUCT.
4. Short-circuit `IF`/`IFERROR`; AND/OR/NOT.
5. VLOOKUP, MATCH, INDEX and nested function calls.
6. SUMIF, COUNTIF and AVERAGEIF criteria.
7. Numeric rounding, logarithmic and text functions.
8. Syntax errors, caret offsets, division by zero, missing references and bad argument errors.
9. Direct and indirect circular references.
10. Selective recalculation and graph edge replacement.
11. Transactional undo/redo and `$`-locked drag-fill shifts.
12. What-if scenario isolation.
13. JSON workbook round-trip and invalid imports.
14. CSV quoting and invalid CSV recovery.
15. 12,000 populated cells with only the touched cell recalculated on an isolated edit.

The test runner returns a nonzero exit code on failure. Do not submit results from this document as evidence unless the tests have been run on the submitted commit.

**V2 regression tests (16 new tests):** number formatting, safe styling validation, formatting undo/redo, combined edit/style atomic undo, dependency cache invariance, independent/duplicated sheet data, sheet validation/deletion, JSON v1-to-v2 migration, find/replace with formula protection, numeric and formula-aware sorting, and data-quality selection analytics. The observed local suite was **42/42 passing**.

### Optional compiled-browser smoke test

The package also includes `scripts/browser_smoke_optional.py` for developers who have Python and Playwright installed. It assembles actual `dist/` JS modules through a browser import map and tests grid editing, undo/redo, formatting, clipboard formulas, worksheet navigation, find, Analyze, sorting, and a v2 JSON export. This bypasses direct HTTP navigation, which the original development environment blocked. **It does not test localStorage, actual server hosting, or GitHub Pages**; check these in a real browser manually.

```bash
python -m pip install playwright
python -m playwright install chromium
npm run build
python scripts/browser_smoke_optional.py
```

If your machine already has a Chrome/Chromium executable, set `GRIDFORGE_CHROMIUM` to its full path rather than installing a Playwright-managed browser. Playwright is optional; it is not required to use or build GridForge.

## 3. Manual acceptance matrix

| Test | Steps | Expected |
| --- | --- | --- |
| Initial state | Start app, inspect D4 and total row | D4 = 4000; D14 = 34739; grid responds |
| Keyboard editing | Go to B4, type `225`, Enter | B4 = 225; D4 = 7200 |
| Formula inspector | Select D4 | AST shows multiplication of B4 and C4 |
| Dependency impact | Select B4 → Impact | Shows D4 and other downstream cells |
| What-if isolation | Select B4 → What-if, preview `500` | Diff appears; B4 stays unchanged until Apply |
| Syntax diagnostic | Enter `=SUM(A1, B2` in I25 | #PARSE!, missing `)`, caret at end |
| Direct cycle | Put `=A30+1` in A30 | #CYCLE!, no browser freeze |
| Indirect cycle | A30=`=B30+1`; B30=`=A30+1` | Both display #CYCLE! |
| Drag-to-fill | H4=`=B4*2`; drag handle down through H7 | H5=`=B5*2`; H7=`=B7*2` |
| Undo grouped fill | Ctrl+Z once after drag fill | Fill operation undone as a single transaction |
| New workbook | Click New and confirm | Empty 400×52 sheet |
| JSON | Export JSON, modify cell, reimport | Original formula source and values restored |
| CSV | Export sample to CSV | Evaluated totals rather than formula strings |
| Scroll | Jump to AZ400 | Grid renders quickly without mounting all 20,800 cells |
| Reload | Edit a cell, refresh | Changes persist when browser storage is enabled |
| Resize | Reduce viewport to tablet width | Sheet remains navigable; sidebar remains reachable |
| Worksheet add/duplicate/rename | Create Sheet 2, type 42 in A1, return to Sheet 1, duplicate | Independent formulas, original unaffected |
| Formatting | Set number format to Percent on cell with numeric `0.125`; then undo | Shows `12.5%`, undo returns general formatting |
| Find/Replace | Ctrl+F, search P-103, replace a text cell | Finds clickable refs, formulas protected by default |
| Data Quality Lens | Select A4:B11 then Analyze | Statistics and bar chart update for selected range |
| Sorting | Select A4:B11 then Sort A–Z; undo | SKU/Unit pairs reorder then restore |
| Same-tab clipboard | Copy formula with Ctrl+C; paste elsewhere in grid | Relative references shift, formatting preserved when custom MIME works |
| Formula functions | Check SUM, IF, VLOOKUP | All behave as documented |
| Browser console | Repeat core edits | No unhandled JavaScript exceptions |

## 4. Short demo-video storyboard (~3–4 minutes)

**0:00–0:15** Introduce GridForge as an explainable custom TypeScript spreadsheet with 20,800 virtualized cells and 136 functions.

**0:15–0:40** Enter `225` into B4. Show D4, F4, and total formulas react immediately. Point out recalculated cell count.

**0:40–1:00** Show D4's expression tree in Inspector and direct references. Open Impact for B4.

**1:00–1:25** Preview a hypothetical change to B4 in What-if. Explain why the original workbook stays unchanged.

**1:25–1:50** Demonstrate a syntax error (`=SUM(A1,B2`) with a caret suggestion and a circular reference.

**1:50–2:15** Create a formula with relative references, drag-fill, undo, redo and demonstrate keyboard navigation.

**2:15–2:40** Add a sheet, duplicate it, and demonstrate bold/currency formatting and undo.

**2:40–3:00** Select a range, sort and undo; use Find to locate P-103.

**3:00–3:20** Open Analyze for A4:B11; show live selection statistics and chart.

**3:20–3:40** Export and import a v2 JSON document, show the recorded benchmark, and close with GitHub Actions status.

## 5. User-explainable interview answers

**Why build your own formula parser?** The brief explicitly requires custom formula logic. A tokenizer and Pratt parser make precedence and nested expressions explicit and testable without `eval()`.

**How do you avoid recomputing the whole sheet?** Every formula records its upstream references and the reverse graph maps each input to downstream formulas. A changed cell starts a graph traversal that invalidates only reachable dependents. Cached unrelated cells are left alone.

**How are circular references detected?** Each evaluation carries its active recursion path. Revisiting a reference already in that path signals a cycle, yielding an error with the path instead of looping forever.

**What is virtualization?** The scroll surface pretends to occupy the full 20,800-cell address space, but only row and column indices intersecting the viewport plus a small buffer are turned into DOM nodes.

**How does What-if prevent data damage?** It clones the raw workbook, applies the hypothetical change to the clone, recalculates there and compares values. The real workbook changes only if the user presses Apply.

**Why no React?** The framework is optional, and the required engineering challenge is the spreadsheet engine. Vanilla TypeScript keeps the data flow explicit and removes runtime dependencies. The engine is framework-agnostic and can be reused in React later.

## 6. Release checklist

- [ ] `npm run check` passes.
- [ ] `npm test` passes.
- [ ] Manual browser acceptance checks pass on the submitted commit.
- [ ] `README.md`, architecture, v2 feature guide, function reference and v2 screenshot are present.
- [ ] `dist/` reflects the latest sources after `npm run build`.
- [ ] Verify no local data, credentials, environment secrets, or OS-specific absolute paths are committed.
- [ ] Push all files except `node_modules` and ignored temporary assets to a public GitHub repository.
- [ ] Only publish a deployment URL after deploying and verifying it.
- [ ] Record a short demo showing source + browser and failure cases, not just UI polish.
