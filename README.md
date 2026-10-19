# GridForge v2.0 · Spreadsheet Intelligence

![GridForge spreadsheet interface](docs/preview_v2.png)

**Data Quality Lens:** [see selected-range analytics screenshot](docs/preview_analytics_v2.png).

A browser-based, Excel-inspired spreadsheet built in **TypeScript** for **GDG Tech Wildcard Entry · Frontend Task 2**. The expression tokenizer, Pratt parser, custom 56-function evaluator, dependency graph, circular-reference detection, and incremental recalculation remain independent of any third-party formula engine.

**Version 2 adds:** independent worksheet tabs, style/number formatting, sortable data ranges, find and replace, selection analytics with data-quality checks and a quick chart, formula-aware same-app copying, values-only paste, per-sheet undo/redo, and one-click GitHub Pages deployment workflows.

## Quick start

Prerequisite: **Node.js 20 or newer** and npm (Node 22 is used in GitHub Actions).

```bash
npm install
npm run dev
```

Visit **http://localhost:4173**. On Windows, after extraction, you may instead double-click `run_windows.bat` to serve the **included prebuilt `dist/`** using Node, without an npm install. On macOS/Linux run `./run_unix.sh` (or `bash run_unix.sh`). No backend, database, API token or paid cloud account is required.

## Excel-like tools

| Feature | How to use it | Scope |
| --- | --- | --- |
| Worksheets | Add `+`, click a tab, Rename, Duplicate, Delete | Up to 12 independent sheets |
| Formatting | Select cells; use **B**, *I*, underline, font/fill colors and alignment | Sparse per-cell styles; JSON-persisted |
| Number formats | General, two decimals, integer, USD, percent | Affects display, never stored numeric precision |
| Grid editing | Type, double-click, Enter/F2, arrows, Shift+arrows | 400 × 52 per sheet by default |
| Formula bar | Edit raw formula such as `=SUM(B4:B11)` | 56 custom functions |
| Drag-to-fill | Drag selected cell's bottom-right handle | Relative/absolute reference shifts |
| Undo/redo | Ctrl+Z, Ctrl+Y, toolbar buttons | Edit and formatting transactions; per sheet, current session |
| Copy/paste | Ctrl+C, Ctrl+V while grid focused | Same-app raw formulas + formatting, external TSV text |
| Paste values | Copy a range; select destination; click `Paste values` | Calculated values, without formulas |
| Sort | Select a rectangular range of ≥2 rows; Sort A–Z or Z–A | Entire selected rows sorted on leftmost selected column |
| Find/replace | Ctrl+F or `Find` | Active sheet; replace excludes formulas by default |
| Analyze | Select range; open Analyze sidebar tab | Count/sum/average/min/max/blanks/errors/duplicates + chart |
| Inspector | Click cell; open Inspector | AST, inputs, parsed formula errors, suggestions |
| Impact | Click cell; open Impact | Upstream/downstream dependency visualization |
| What-if | Select cell; enter alternate content; Preview | Non-destructive simulation with differences |
| Persistence | Autosaves in your browser | All sheets and formatting via localStorage |
| Import/export | JSON, CSV | JSON retains formulas/sheets/styles; CSV exports active sheet values |

**Examples:** `=SUM(B4:B11)`, `=IF(D4>5000,"High","Normal")`, `=VLOOKUP(J5,A4:D11,4,FALSE)`, `=ROUND(F4*0.18,2)`.

## Engineering proof

- Grid space: **20,800 addressable cells per standard sheet**; viewport virtualization renders only visible cells.
- Formula library: **56 hand-implemented functions** with documented supported semantics.
- Dependency graph: forward and reverse edges; edits invalidate transitive downstream cells only; direct and indirect cycles report `#CYCLE!`.
- Error diagnostics: parser caret position and suggested correction; Inspector surfaces details.
- Unit tests: **42 passing** at the v2 checkpoint (26 original + 16 new), run using Node's built-in test runner.
- Reproducible benchmark: **10,000 formulas / 10,400 populated cells**; on the recorded local run, an input change recalculated just **26 cells**. Hardware-dependent time measurements are in `docs/PERFORMANCE_REPORT.md`.

```bash
npm run check        # TypeScript type check
npm test             # builds and runs automated tests
npm run benchmark    # independent engine performance experiment
npm run build        # emits static dist/ for deployment
npm start            # serves dist/ at port 4173
```

## GitHub repository and deployment

This ZIP is structured as the **repository root**, not nested inside an extra generated parent folder. After extraction:

```bash
git init
git add .
git commit -m "GridForge v2: multi-sheet spreadsheet with custom formula engine"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPOSITORY.git
git push -u origin main
```

Create an **empty GitHub repository first**, and replace the placeholder with its real URL. **Do not paste the placeholder literally.** GitHub Actions workflows included:

- `.github/workflows/ci.yml`: install, type check, full tests, benchmark, compiled-site artifact.
- `.github/workflows/pages.yml`: rebuild, test, and deploy the `dist/` site to GitHub Pages on pushes to `main`.

In the repository **Settings → Pages**, choose **GitHub Actions** as the build/deployment source. See `docs/GITHUB_RELEASE.md` for troubleshooting and reviewer demonstration steps. A public GitHub URL is not included until you create and push the actual repository.

## Important compatibility limits

GridForge is **Excel-inspired, not an Excel clone**. The spreadsheet syntax intentionally uses **sheet-local references** (`A1`, `B2:C5`); **cross-sheet formulas** (`Sheet2!A1`) are not implemented. **`.xlsx` import/export is not implemented**; use CSV for active-sheet values or the app's JSON format to preserve formulas, tabs, and formatting. Sorting shifts relative references inside formulas moved to new rows; this is a documented simplified sorting policy, not a claim of identical Excel behavior. The chart is a lightweight sidebar visualization, not a saved Excel chart object. There are no pivots, shared collaboration, macros, charts persisted in the workbook, or full Excel date semantics. State is browser-local and subject to storage quotas; export JSON before clearing browser data.

## Project structure

```text
.github/workflows/   CI and GitHub Pages workflows
src/engine/parser.ts        Tokenizer, Pratt parser, AST, relative references
src/engine/functions.ts     56 custom spreadsheet functions
src/engine/workbook.ts      Cells, cache, dependencies, transaction undo/redo
src/engine/format.ts        Safe cell styles + number formatting
src/engine/document.ts      Multiple worksheets + v1 import migration
src/engine/operations.ts    Range sorting, searching, analytics
src/engine/csv.ts           CSV parser + serializer
src/main.ts                 Virtual grid, toolbar, dialogs, clipboard, UI
public/                    HTML + CSS
dist/                      Precompiled static site (versioned in ZIP)
scripts/                   Serve/build/benchmark scripts
tests/                     42 independent automated tests
docs/                      Technical and reviewer documentation
run_windows.bat / run_unix.sh
```

## Documentation and reviewer evidence

- [Architecture](docs/ARCHITECTURE.md) · [Full list of 56 functions](docs/FUNCTION_REFERENCE.md)
- [v2 feature design and behaviors](docs/V2_EXCEL_FEATURES.md) · [User guide](docs/USER_GUIDE.md)
- [Test and live demo walkthrough](docs/TESTING_AND_DEMO.md) · [Performance evidence](docs/PERFORMANCE_REPORT.md)
- [GitHub upload and Pages deployment](docs/GITHUB_RELEASE.md) · [Project continuation handover](docs/PROJECT_HANDOVER.md)

MIT licensed. All spreadsheet computation runs entirely in the browser. The included demo workbook is synthetic.
