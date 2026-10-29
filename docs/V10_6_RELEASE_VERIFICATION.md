# GridForge v10.6 — Release Verification and Scope

## Verified in this environment

- TypeScript compiler: PASS (`tsc -p tsconfig.json`).
- All existing and new Node unit tests: **847 PASS, 0 FAIL** (`node --test tests/*.test.mjs`). Includes seven new v4 executor tests.
- Earlier offline macro v1/v2/v3 compatibility tests: PASS.
- Offline CLI preview with JSON workbook and parameter overrides: PASS.
- Offline CLI apply and new GridForge JSON saved: PASS.
- Python-only packaged app asset check: PASS (`python run_gridforge.py --check`).
- Desktop wrapper Python unit tests: **7 PASS**.
- 10,000-formula dependency benchmark: PASS (single source edit invalidated 26 cells; environment-specific timings not treated as Windows performance claims).
- ZIP integrity and file inventory: checked at packaging.

## Not verified here

- Interactive Chromium/WebView2 automation: hosted browser navigation blocked by administrator policy, even for localhost/intercepted origin. Use `python tests/smoke_macro_studio.py` on Windows.
- Actual Windows PyInstaller packaging and Inno Setup installer: must run `BUILD_WINDOWS_INSTALLER.bat` on a Windows x64 machine with prerequisites.
- Existing Excel VBA execution via Microsoft Excel COM: not tested here; the browser-native macro engine needs no Excel and never executes embedded VBA.
- Complete Microsoft Excel behavioral equivalence is not claimed; the new Excel-like cleanup/format actions implement a documented subset.

## Acceptance checklist for Windows distribution

1. Build `GridForge-Setup-10.6.0-win64.exe` on Windows.
2. Install on a clean Windows account and verify both Start menu shortcuts.
3. Launch **GridForge Macro Studio** directly and confirm the Automate dialog opens.
4. Import `examples/macros/parameterized_tax.json`, override `tax`, preview and run, verify cell values and undo.
5. Import `examples/macros/cleanup_contacts.json`, preview text edits and duplicate compaction before applying.
6. Export and re-import the macro definition JSON; export the execution report.
7. Save and reopen XLSX/XLSM, verify macro provenance and offline behavior.
8. Validate uninstaller; check that it does not delete user drafts without warning.
9. Verify installer file hash, malware scan and consider code signing before public release.

## Limitations

`removeDuplicates` compacts the selected cell rectangle (not complete spreadsheet row deletion), `formatIf` changes static cell styling rather than registering live conditional-format rules, and `cleanText` skips formula source by default. Neither GridForge native macros nor the CLI executes arbitrary VBA or JavaScript.
