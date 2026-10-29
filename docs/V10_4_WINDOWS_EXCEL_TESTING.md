# Windows — Actual Excel check for the ten connected names

This test was **not run on the authoring Linux environment** because Excel desktop COM is Windows-only. It is *not* a test of successful cube, PivotTable, market data or real-time integrations. It tests how missing connections are reported when the same formulas are evaluated in a blank Microsoft Excel workbook.

Prerequisites: Windows, Microsoft Excel desktop (licensed), Python and `pywin32` installed into the same active interpreter.

```powershell
python -c "import win32com.client; app=win32com.client.DispatchEx('Excel.Application'); print(app.Version); app.Quit()"
python scripts\excel_connected_windows.py
```

Alternative: double-click `RUN_CONNECTED_PARITY_WINDOWS.bat`.

The runner creates a new temporary unsaved workbook in a separate Excel COM process. It does **not** read your files or execute VBA. It avoids setting Excel's `Application.Calculation` property, which previously caused a COM failure on your computer.

Expected output: `docs/V10_4_CONNECTED_EXCEL_WINDOWS_RESULTS.json` with individual `MATCH_OFFLINE_ERROR`, `DIFFERENT_OFFLINE_ERROR`, `COM_OR_ASSIGNMENT_ERROR`, or `SKIPPED_EXTERNAL` statuses. Missing-connection matches are not evidence that a successful connected calculation would match Excel. RTD, market history and web requests are skipped to prevent unintended external access.

Excel versions can differ in displayed errors, function availability and calculation semantics. Keep exact formulas, `excelVersion`, `excelBuild` and raw errors in the report. Avoid altering GridForge's errors solely to inflate a match percentage without validating behavior.

The earlier 506 representative-function parity report remains a distinct evaluation of the **506 locally implemented names** (not these ten). Refer to `docs/V10_3_EXCEL_REPAIR_AUDIT.md` for the frozen v10.3 snapshot comparison, which was **not rerun** in v10.4.
