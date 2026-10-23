# GridForge v6: reviewer demonstration and acceptance checks

1. **Python-only launch:** In an extracted copy, run `python run_gridforge.py`, visit the printed URL, and verify a demo table appears. No third-party Python packages are required.
2. **Excel-style category browser:** Open the Functions sidebar, confirm it shows **306 built-ins**, choose **Recently added** (90 entries), and filter by Statistical for `STDEV.S` and `NORM.DIST`.
3. **Dotted formulas:** Enter `=NORM.S.DIST(0,TRUE)` (0.5), `=BITAND(6,3)` (2), and `=IFNA(NA(),77)` (77).
4. **Runtime dependencies:** Enter `8` in A1, `=INDIRECT("A1")*2` in B1, change A1 to `12`, and confirm B1 changes from 16 to 24.
5. **Gallery:** Import `examples/v6_90_function_gallery.xlsx` and inspect the `90 Functions v6` sheet. One row per new function is evaluated; `NA()` intentionally shows `#N/A`.
6. **Recorder:** Choose Automate → Start recording, edit a cell, apply Bold, switch sheets and edit another cell, stop and save. Preview effects and confirm no changes are committed until Run is selected.
7. **Advanced JSON macro:** On a workbook with `Sheet 1`, import `examples/v6_safe_multisheet_macro.json`, preview, run and undo its changes. Test bad JSON with an unknown step kind: it must be rejected.
8. **XLSM safeguard:** If using a test `.xlsm`, ensure the original download matches the imported file byte-for-byte. Ensure exporting edited `.xlsx` creates a VBA-free file. Do not expect embedded VBA to execute.
9. **Regression:** `npm run regression` after installing Node.js, Python Playwright and Chromium. Expected release baseline: 340 Node tests, 24 browser checks, four stages.
10. **Local browser persistence:** Close GridForge's tab and reopen the **same origin**; check that the edited workbook restores. This check was skipped in the hosted test environment and should be performed on Windows.

Record actual results and failures. Do not claim the test passes if a reviewer observes a mismatch.
