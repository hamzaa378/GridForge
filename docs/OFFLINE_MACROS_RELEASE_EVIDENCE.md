# Offline macro enhancement test evidence (v10.5 source)

Tested on the packaged TypeScript source and generated JavaScript frontend:

- TypeScript compile: PASS.
- `npm test`: 840/840 PASS. Existing 834 plus six new v3 macro engine tests.
- `node --test tests/macros_v6.test.mjs tests/macros_v3_offline.test.mjs`: 11/11 PASS.
- Desktop Python unit tests: 5/5 PASS.
- `python run_gridforge.py --check`: PASS.
- New `ifCell`, `sequence`, `repeat` routines tested for calculation, rollback on error, limits and v1/v2 compatibility.
- Browser UI smoke: **NOT VERIFIED HERE**. Hosted Chromium reports `ERR_BLOCKED_BY_ADMINISTRATOR` navigating to localhost. Run the Windows acceptance checklist below.
- Windows `.exe` and Inno Setup installer: **NOT BUILT HERE**; build using `BUILD_WINDOWS_INSTALLER.bat` or the supplied GitHub Actions Windows workflow.
- Imported Excel VBA execution without desktop Excel: **NOT SUPPORTED**. No full VBA runtime is embedded.

## Required Windows acceptance

1. Build the portable and installer executables, or start `python run_gridforge.py` for browser-only UI testing.
2. Open **Automate** and create the **Number rows 1–20** macro template.
3. Preview it. Verify A2:A21 do not change before execution.
4. Run it. Verify A2:A21 contain 1–20. Undo.
5. Import `examples/native_macros/conditional_review.gfmacro.json` and test B2 above and below 100.
6. Export the macro and import it into a new workbook with matching worksheet names.
7. Close and reopen the desktop application; verify the macro library persists.
8. Verify imported `.xlsm` files do not run VBA without explicit optional Excel bridge approval.
9. Run `RUN_OFFLINE_MACRO_TESTS_WINDOWS.bat` on a Node-enabled builder.
10. Test installer, portable GUI, saving workbooks, and uninstall on a clean Windows machine.
