> **GridForge v10 note:** The current release has 506 registered spreadsheet functions, a persistent Light/Dark switch and Excel familiar/Guided workspace modes. Full v10 behavior and test evidence are in [the v10 release report](V10_RELEASE_EVIDENCE.md). When running with Python alone, use the included compiled `dist/`; Node.js is only needed to rebuild source. Some older sections below retain their historical version numbers.

# GridForge v5.0.0 — regression testing and failure guide

The current v5 release adds another 80 spreadsheet functions (216 total) and expands automated coverage to 241 unit tests plus 21 browser checks. This guide documents how to detect regressions *before* a push, not a guarantee that the application is defect-free.

## Setup: Windows PowerShell

Extract the repository. Install Node.js 20+ and Python 3.10+ and run these commands from the project root:

```powershell
npm install
py -3 -m pip install playwright
py -3 -m playwright install chromium
npm run regression
```

If `py` is unavailable, use `python` in its place. For a simpler PowerShell entry point, run `./run_regression_windows.ps1` (PowerShell execution policy may need adjustment). **You do not need to start `npm run dev`**: the browser test launches and stops its own server on an available random port. It will not occupy port 4174, modify your existing browser profile, or overwrite your saved GridForge workbook.

## Test commands

| Command | What it verifies |
|---|---|
| `npm run regression` | Typecheck, complete unit tests, dependency benchmark, Chromium UI suite |
| `npm run regression:core` | The first three stages only; useful if Chromium not yet installed |
| `npm run regression:browser` | The Chromium UI suite only (requires `npm run build` first) |
| `npm test` | Build and 241 Node.js test cases |
| `npm run check` | TypeScript type checking |
| `npm run benchmark` | The 10,000-formula graph must recalculate only 26 cells for the dependent edit and one for an unrelated edit |

All commands return a nonzero exit code on a failure. The combined runner stops at the *first failing stage*, prints the command and error, and reports how many stages passed.

## What the 241 Node tests cover

The **original 65 tests** cover formula parsing/evaluation, the earlier 56-function baseline, error diagnostics and cycles, undo/redo, JSON, copy/paste, formatting, multi-sheet documents, custom sorting, XLSX writer structures, filters, and 12,000-cell sampling. **91 v4 test cases** cover the first 80 new functions and their edge cases; **85 v5 cases** cover another 80 functions, Excel-style categories, macro validation, XLSX compatibility and security checks. Earlier regression canaries include:

1. Repeated input changes update only the dependent formula chain and never unrelated cells.
2. Editing after an Undo discards the obsolete Redo history.
3. A malformed style transaction cannot partially modify the workbook.
4. Sorting numeric rows keeps their formulas, formatting and full Undo snapshots coherent.
5. Breaking a circular dependency clears former `#CYCLE!` results.
6. Copying relative and absolute references keeps locked rows and columns fixed.
7. Restored worksheets have independent cell values, caches and styles.
8. XLSX export contains essential Office ZIP parts and actual formula XML.

## What the 21 browser regression checks cover

The `scripts/regression_browser.py` script uses a new headless Chromium browser profile and an ephemeral HTTP port to test the actual compiled UI:

- Demo workbook loading and recalculated formula results.
- Cell editing, Undo and Redo, and formula parser error feedback.
- Bold formatting, Undo/Redo for styles.
- Custom right-click menu, copy/paste with relative formula shifts, keyboard context menu.
- Custom sorting with header row, numeric sort, aligned rows, and undo.
- Find and replace with undo.
- Conditional formatting and a temporary filter, then restoring all rows.
- Freeze row and column; AutoSum generates the expected formula.
- Create and duplicate worksheets; check independent data.
- JSON export and real localStorage recovery after a page reload.
- Genuine `.xlsx` download, ZIP integrity, import of the downloaded workbook and independent sample fixture import.
- Zero uncaught JavaScript exceptions.

**Hosted-browser limitation:** Some managed environments block all browser navigation with `ERR_BLOCKED_BY_ADMINISTRATOR`. In that case the test automatically exercises an in-memory page containing the same compiled modules. It **clearly reports that localStorage reload was skipped**; do not treat this as a fully complete browser-storage test. On a normal Windows machine, the test uses a real HTTP origin and checks storage reload.

## Troubleshooting

| Symptom | Action |
|---|---|
| `py` or `python` not found | Install Python, reopen PowerShell and retry |
| `No module named playwright` | Run `py -3 -m pip install playwright` |
| Chromium executable missing | Run `py -3 -m playwright install chromium` |
| Browser cannot launch | Verify Chromium install; optionally set `$env:GRIDFORGE_CHROMIUM='C:\path\to\chrome.exe'` |
| Server port conflict | Runner requests a free ephemeral port; close unusual software reserving all ports and retry |
| `npm run check` fails | Fix the TypeScript compiler error, then rerun tests |
| Unit test fails | Read the failed test's expected and actual values. Do **not** delete the test to make CI green |
| XLSX test fails | Check the first reported workbook mismatch; confirm file name ends with `.xlsx` |
| Regression passes locally but CI fails | Check the GitHub Actions logs for differences in Node, Python, browser version, or OS |

## CI and limitations

`.github/workflows/ci.yml` runs `npm run regression` for pushes to `main`, pull requests to `main`, and manual runs. A GitHub check passing means the code met these specific assertions in that environment, not that every Excel file will work.

**Not fully automated:** Real-world VBA compatibility (VBA never executes or merges into edited XLSM), Microsoft Excel/LibreOffice interoperability, large real-world imported datasets, cross-browser/device accessibility, drag interactions on touch screens, screenshots/pixel-perfect visual regressions, offline persistence across different origins, and unsupported Excel features such as VBA execution and pivot tables. Test these manually when relevant. Do not use `.xlsx` import/export as an exact-preservation backup for unsupported Excel features; use JSON for GridForge-specific state.

## Before GitHub submission

1. Run `npm run regression` and confirm every stage passes with no omitted browser-storage check on your machine.
2. Optionally test `examples/sample_sales.xlsx` in actual Microsoft Excel or LibreOffice.
3. Run the acceptance checklist in `docs/V3_ACCEPTANCE_CHECKLIST.md` for any functionality you changed.
4. Commit/push, then verify the GitHub Actions check is green. Include reproducible test evidence instead of claiming untested compatibility.


## v5 additions

`tests/functions_v5.test.mjs` checks each of the new 80 functions plus error cases, categories, macro validation/replay and edited XLSX macro omission. `scripts/regression_browser.py` adds browser checks for the category picker, v5 formulas, record/playback/JSON export, the v5 Excel gallery and a **synthetic** XLSM byte-preservation test. In the packaging environment **241 Node tests** and **21 browser checks** passed; browser localStorage reload was skipped due to hosted navigation restrictions. Real Microsoft Excel VBA execution is deliberately never tested or supported by GridForge.
