# Continued work: Excel compatibility

- Base release: GridForge v10, 506 registered functions, dual light/dark and Excel/Guided UI.
- This audit build: v10.1, fixes to rounding, aggregation argument context, criteria wildcards, and an Excel comparison harness.
- Core source changes: `src/engine/functions.ts` and `src/engine/workbook.ts`.
- New tests: `tests/excel_parity_reference.test.mjs` and generated `tests/excel_parity_cases.json`.
- New scripts: `build_excel_parity_cases.mjs`, `excel_parity_gridforge.mjs`, `excel_parity_desktop.py`, `excel_parity_compare.py`, `excel_parity_compare.mjs`.
- Python-only UI: `python run_gridforge.py`, no pip dependencies required.
- Real Excel parity: Windows Excel + `pywin32`, using `RUN_EXCEL_PARITY_WINDOWS.bat`.
- Security: Excel runner generates a fresh unsaved workbook, disables events and macro automation, and only evaluates the included fixture formulas. It does not open external user files.
- Existing VBA execution: remains opt-in through separate `scripts/excel_vba_bridge.py`, requires Excel desktop, never executed in this host.
- Evidence: `EXCEL_PARITY_RELEASE_EVIDENCE.md`. No measured Excel agreement percentage yet.

## Next priorities

1. Run the 140 probes with actual Microsoft Excel. Store the version/build and full results, triage disagreements, then rerun.
2. Add 10–50 real-user scenario workbooks (finance, sales, dates, lookups, formatting, charts) with manual before/after expectations.
3. Implement dynamic-array spilling safely, with cell occupancy checks, `#SPILL!`, and dependency propagation.
4. Add cross-sheet formulas, structured references and named ranges if they are needed for real workbook compatibility.
5. Fuzz argument-count, locale/text/number coercion, financial date conventions and 1900-date edge cases.
6. Verify XLSX and XLSM with **actual desktop Excel**, including recalculation cache, complex formulas, workbook structure and intentionally restricted macros.
7. Repeat unit, browser, benchmark, XLSM and VBA security tests after each change; maintain an honest limitations matrix.
