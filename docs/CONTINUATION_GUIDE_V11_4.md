# GridForge v11.4 — Continuation Guide

GridForge is a standalone spreadsheet application with its own worksheet grid and workbook engine. Excel 2019 is an interaction reference only. The desktop application does not embed or require Microsoft Excel.

## Implemented in v11.4

- Queued worksheet chart rendering so that each visible SVG chart is painted after the grid, with a yield between chart objects.
- Invalidated queued chart work when a newer worksheet render begins.
- Added horizontal axis title editing to the Format ribbon and kept the setting in the chart model.
- Changed the desktop server to use a per-launch ephemeral loopback port.
- Updated desktop version metadata to 11.4.0.

## Development and release

Run `npm run build`, `npm run regression`, `python -m unittest discover -s tests -p "test_desktop*.py"`, and `python scripts/smoke_migration_v110.py`. On Windows, build the portable app and installer with `desktop/build_windows.ps1` or `BUILD_WINDOWS_INSTALLER.bat`.

## Next areas

Continue improving supported Excel chart layouts, keyboard navigation, accessibility, and workbook compatibility. Keep each new format behavior tied to explicit import/export verification and document unsupported Excel features rather than implying full parity.
