# GridForge v8 — Reviewer acceptance checklist

Use this on Windows or Linux/macOS in a fresh browser profile, preferably at `http://127.0.0.1:4174`. Do not import private workbooks into a public demo.

1. Open app using **Python only**: `python run_gridforge.py`. Confirm demo loads and revenue formulas evaluate.
2. Import `examples/v8_data_studio_demo.json`. Select A5:A9, open **Data Studio**, filter **Text & case**, choose **Title Case**, preview, apply, and confirm the names change. Undo/redo.
3. Select B5:B9, choose **Increase 10%**, preview, apply, then undo and confirm the original numbers return.
4. Select a blank rectangle, choose **Series 1, 2, 3…** and confirm sequential numbers filled across rows. Undo.
5. Select a square area, choose **Transpose square**, preview and undo. Select a non-square area; confirm an explanatory error rather than partial changes.
6. Choose **SUM below selection** for B5:B9. Confirm the resulting formula uses `=SUM(B5:B9)`. If the cell below is occupied, a confirmation prompt must appear.
7. Put `=B5*2` in an empty cell and fill downward. Verify relative references change; verify `$`-locked references remain fixed.
8. Save and import a JSON workbook; restart browser and verify data retained. This must be checked on a machine where browser navigation/storage works normally.
9. Export an XLSX and re-import it; verify supported raw formulas and sheet contents.
10. If using an XLSM test fixture with trusted source, edit an ordinary cell and use **Save XLSM**; verify the edited cell and original `vbaProject.bin` are retained. Never execute an untrusted macro.
11. Open **Automate**, preview/replay a safe JSON macro. With Windows Excel installed, optionally follow the separate v7 VBA bridge test instructions (not required for a Python-only demo).
12. Developer machine: run `npm run regression` and confirm all stages pass with no skipped stages other than documented environment limitations.

The reviewer should also check keyboard navigation, right-click Data Studio, undo history, find/replace, sorting, conditional formatting, filters, freeze panes, virtualized scrolling, formula diagnostics and impact analysis. Report any failing behavior with reproducible input, expected output and observed output.
