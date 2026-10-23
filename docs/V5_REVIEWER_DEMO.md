# GridForge v5 — Reviewer demo and acceptance checklist

## Start from a computer with Python and a browser

1. Extract the entire ZIP, keeping `dist/` beside `run_gridforge.py`.
2. In Windows Explorer double-click `START_HERE_WINDOWS.bat`, or from a terminal use `python run_gridforge.py`.
3. Open the printed URL (usually `http://127.0.0.1:4174`). No Node.js or pip packages are needed to **run** the compiled app.
4. Edit a cell in the demo worksheet and verify dependent formulas recalculate. Try Undo and Redo.

## 216 functions in Excel-style groups

1. Open **Functions** in the right sidebar and filter to **Financial**, **Date & Time**, **Lookup & Reference**, or **Math & Engineering**.
2. Search for `XLOOKUP`, `IRR`, `NETWORKDAYS`, `REGEXTEST`, and `MROUND`; verify examples appear.
3. Enter `=MROUND(17,5)` in a blank cell and expect 15.
4. Import `examples/v5_80_function_gallery.xlsx` through the XLSX import control. Inspect representative results for the new v5 functions. This tests GridForge's own parser and evaluator, not an Excel formula library.

## Record and play a GridForge action macro

1. Open **Automate** and select **Start recording**.
2. Enter a value in a blank cell and change its formatting.
3. Return to **Automate**, save the recording under a unique name, and select **Run**.
4. Check the replayed changes, and use Undo. Export the recording as JSON if desired.
5. Explain to evaluators that GridForge action macros can replay **recorded cell and style changes only**. They never execute arbitrary scripts or VBA. Sorting and other actions may not be recorded.

## Import a macro-enabled Excel workbook safely

1. Use an `.xlsm` workbook from a trusted source. Keep Excel macros **disabled** when examining it in Microsoft Excel.
2. Import `.xlsm` and verify supported worksheet values and formulas appear.
3. Select **Original XLSM** to download the **unchanged source file** during that session; check it against the original with a hash if desired.
4. Change a worksheet value in GridForge and select **XLSX** to export updated supported data.
5. **Do not expect the new XLSX to contain VBA or the original XLSM to contain your GridForge edits.** GridForge cannot run, edit, migrate, or reinsert Microsoft VBA. Excel's own runtime is required to run VBA macros safely in a trusted environment.

## Full regression (developer/reviewer with Node.js)

Install Node.js 20+ and Python Playwright with Chromium as described in `docs/SETUP_FROM_ZERO.md`, then run `npm install` and `npm run regression`.

The build-time verification was **241/241 Node tests, 21/21 browser checks, 4/4 stages**, with browser storage reload **skipped** by the hosted testing environment; repeat that step on a normal Windows browser. Real-world Excel and VBA compatibility is not certified.
