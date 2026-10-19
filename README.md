# GridForge v3.0 — Spreadsheet Intelligence

![GridForge v3 spreadsheet with Excel-style grid and frozen panes](docs/preview_v3.png)

**GDG Tech Wildcard Entry — Frontend Task 2.** A browser-based, virtualized spreadsheet written in TypeScript. Features a **handwritten lexer, Pratt parser, evaluator with 56 functions, incremental dependency graph and cyclic-formula diagnostics**. GridForge v3 adds native `.xlsx` read/write, filtered tables, freeze panes, live conditional formatting and AutoSum. **No existing spreadsheet formula engine is used.**

## Quick start: Windows / macOS / Linux

Install Node.js **20+**, unzip the GitHub package, then run in its directory:

```bash
npm install
npm run dev
```

Open the printed address. The default port is **4174**, with fallback to 4175–4184 when occupied. In PowerShell you may set an explicit port:

```powershell
$env:PORT = "4190"
npm run dev
```

Windows users can run `run_windows.bat` to serve the **included precompiled `dist/`** with Node; it does not need `npm install`. For Unix use `bash run_unix.sh`. GridForge is a static web app, **no backend/database/cloud account required**. Data autosaves per browser origin in `localStorage`; export JSON/XLSX for portable backups.

## Excel features now implemented

| Feature | How to use | Important scope |
| --- | --- | --- |
| **XLSX export** | **XLSX** in the upper toolbar | All sheets, formulas, cell styles and supported table view/rules; Office Open XML workbook |
| **XLSX import** | **Import** and choose `.xlsx` | Standard, unencrypted OOXML; shared strings, inline strings, basic formulas/styles, supported freeze/filter/conditional formatting; review import report |
| **Filter** | Select table including header, click **Filter**, choose column and substring | Non-destructive; filters a single column; **Show all** restores all rows |
| **Freeze panes** | **Freeze row**, **Freeze col** in formatting ribbon | Top row and/or first column remain visible while scrolling; click again to unfreeze |
| **Conditional formatting** | Select numbers → **Conditional** → choose >, < or =, threshold, color | Live numeric comparison; up to 30 rules per sheet; persists in JSON and XLSX |
| **AutoSum** | Click total cell below contiguous numbers, then **∑ AutoSum** or `Alt+=` | Inserts a real `SUM` formula. Selecting a vertical range inserts below selection |
| Custom Sort | Select entire table, right-click → **Custom sort** | Choose column, ascending/descending, and whether the selection has a header; formats move with rows |
| Find & Replace | `Ctrl+F` | Case-sensitive option; formula source protected by default |
| Right-click | Mouse right-click or `Shift+F10` | Copy, cut, paste, values-only, formatting, sorting, Analyze, Inspector, AutoSum, Filter |
| More common tools | Formatting ribbon | Bold/italic/underline, text/fill color, alignment, numeric/currency/percent formats, undo/redo |
| Multiple sheets | Tabs at bottom | Up to 12 independent tabs; create, switch, rename, duplicate, delete |
| Import/export fallback | CSV and JSON | CSV exports values; JSON round-trips GridForge workbook with formulas, styles, filters, freeze and rules |

The application also includes drag-to-fill, formula bar, cell selection, relative/absolute references, VLOOKUP, IF, SUM, conditional aggregations, Find, data-quality analytics, Formula Inspector, Dependency Impact Explorer, What-if previews and keyboard shortcuts.

**Spreadsheet size:** default grid is 400 rows × 52 columns = 20,800 addressable cells per sheet, with viewport rendering (virtualization). Imported sheets may grow up to 1,000 rows × 104 columns; up to 12 sheets.

## Try the included XLSX file

1. Click **Import** → `examples/sample_sales.xlsx`.
2. Close the **XLSX import report**. Check the **Sales** and **Summary** tabs.
3. On Sales, select `A1:D5`. Choose **Filter** → column A → Contains `Mouse`. Only the header and Mouse row remain visible; **Show all** returns every row.
4. Select `D2:D5` → **Conditional** → greater than `100000` → choose color. Changing a revenue formula changes the highlight automatically.
5. Scroll down/right with Freeze row/col active. Click **XLSX** to export, then import the downloaded workbook again.
6. Test `=SUM(B2:B5)`, `=IF(D2>1000,"High","Low")` and the right-click Formula Inspector.

For a complete reproducible acceptance walk-through, see [`docs/V3_ACCEPTANCE_CHECKLIST.md`](docs/V3_ACCEPTANCE_CHECKLIST.md).

## Engineering evidence

| Check | Result in development environment |
| --- | --- |
| TypeScript strict type checking | PASS |
| Automated Node tests | **57/57 passing** (v2.1 had 49) |
| Chromium v3 browser regression | PASS: editing, formatting, multi-sheet, right-click menu, freeze, filtering, conditional formatting, AutoSum, XLSX export/import |
| XLSX round-trip compatibility | PASS: exported stored-ZIP OOXML, recompressed to normal DEFLATE `.xlsx`, imported via file picker; supported formulas, styles, panes and rules verified |
| ZIP archive validation | PASS (CRC and archive part verification) |
| Dependency stress benchmark | 10,000 formulas; edit to one row's source invalidated/recalculated **26 cells**, not 10,000 |

All timing results are environment-dependent; the reproducible script is `scripts/benchmark.mjs`. Browser smoke testing is optional and needs local Python Playwright and Chromium; it is **not** run by standard `npm test` or GitHub Actions. See [`docs/TESTING_AND_DEMO.md`](docs/TESTING_AND_DEMO.md).

```bash
npm run check          # TypeScript static checks
npm test               # builds and runs 57 Node tests
npm run benchmark      # dependency-recalculation benchmark
npm run build          # compiles and copies public files into dist/
npm start              # serves dist/, default http://localhost:4174
node scripts/generate_sample.mjs   # regenerate example XLSX after a build
```

## How XLSX works without a spreadsheet engine

`src/engine/xlsx.ts` creates/parses **Office Open XML** parts (workbook, worksheets, relationships, styles) packaged with `src/engine/zip.ts`. The archive writer uses valid **stored** ZIP entries, the importer reads stored and **DEFLATE** entries using the browser `DecompressionStream('deflate-raw')`. Cell formulas are parsed and evaluated **only** by GridForge's own `parser.ts` and `workbook.ts`. Imported formula syntax that GridForge cannot evaluate is shown as a formula error rather than silently evaluated by a library. Unsupported items are disclosed in the import report.

**Interoperability is a supported subset, not full Microsoft Excel compatibility.** Imported charts, pivot tables, macros, cell merges, rich-text runs, external links, password-protected/ZIP64 workbooks, Excel date semantics, complex filter conditions, advanced conditional rules, named ranges, array/spill formulas and cross-sheet formulas are not supported. Large XLSX sheets above GridForge limits are rejected instead of silently truncating values. **JSON is GridForge's loss-minimizing backup format** for its supported features. Verify imported files before overwriting originals. Original `.xlsx` files are never modified in place.

## Repository layout

```text
.github/workflows/             CI test workflow + GitHub Pages deploy workflow
src/main.ts                    Spreadsheet UI: virtual grid, dialogs, keyboard/mouse interaction
src/engine/parser.ts           Handwritten lexer, parser, AST, reference shifting
src/engine/functions.ts        Custom 56-function library
src/engine/workbook.ts         Sparse cells, dependency graph, cycles, undo/redo
src/engine/document.ts         Multi-sheet container, v2 JSON compatibility, v3 metadata
src/engine/sheettools.ts       Non-destructive filters, freeze/conditional rule state
src/engine/xlsx.ts             OOXML workbook import and export (no formula engine)
src/engine/zip.ts              Minimal CRC-verified ZIP import/export
src/engine/operations.ts       Search, replace, sorting
src/engine/clipboard.ts        Formula-aware cut/copy/paste
public/                        HTML and CSS, including theme/virtualized grid
scripts/                       Static server, build helper, benchmark, optional browser smoke
examples/sample_sales.xlsx     Self-generated, importable sample workbook
tests/                         Node tests for parser, engine, sorting, ZIP/XLSX, views
dist/                         Generated runnable static site (included for Windows launcher)
docs/                          Architecture, release notes, demo, troubleshooting and handover
```

## Push to GitHub and deploy

Create a **new empty GitHub repository** first. From the extracted root (the directory containing `package.json`):

```bash
git init
git add .
git commit -m "GridForge v3: XLSX interoperability and spreadsheet productivity"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPOSITORY.git
git push -u origin main
```

Replace the GitHub URL with the actual empty repository's URL. Both `.github/workflows/ci.yml` and `.github/workflows/pages.yml` are included. After pushing, go to **GitHub → Settings → Pages → Build and deployment → GitHub Actions** and check the Actions logs for the deployment URL. See [`docs/GITHUB_RELEASE.md`](docs/GITHUB_RELEASE.md). Do not upload `node_modules` or any private data.

**Evaluation:** This project satisfies the GDG task's custom-formula-engine rule: the formulas and recalculation logic are authored locally, not delegated to SheetJS, ExcelJS or another spreadsheet evaluator. Be ready to explain parsing, dependency propagation, and v3 XLSX boundaries.

Licensed MIT; see `LICENSE`.
