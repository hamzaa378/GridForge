GridForge v12.0.210 — GitHub-ready source and Windows package

This archive contains the current working-tree source, compiled web frontend (dist/),
tests, documentation, examples, dependency lockfiles, Windows build/installer scripts,
and the compiled Windows installer and portable build. Install dependencies with
`npm ci`; run `npm test`; build the desktop installer on Windows with
`powershell -ExecutionPolicy Bypass -File .\desktop\build_windows.ps1`.

The included installer was rebuilt from the source snapshot in this package. The
Windows packaging build completed successfully. The application UI was not launched
or live-tested. The installer is unsigned. Review docs/DESKTOP_BUILD_AND_RELEASE.md
for prerequisites, manual build steps, and release acceptance guidance.

The bounded Power Query M, Data Model/DAX, connector, PivotTable and chart behavior
is documented in docs/EXCEL_COMPATIBILITY_AUDIT.md and the linked feature notes;
this package does not claim full Excel/Calc parity.
