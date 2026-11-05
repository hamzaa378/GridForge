# GridForge v12.0.209 continuation handoff

## Current work

This working copy advances GridForge from v12.0.208 to v12.0.209. The main change is multi-pair XY scatter charts: up to 24 separate X/Y pairs, editor support, per-series scatter styling, X/Y error bars and trendlines, preview support, and native XLSX/ODS import/export. Hidden pairs are omitted from GridForge preview/editor data and ODS output while remaining series retain their original style indexes. XLSX preserves full native source ranges and uses Excel’s visible-cells-only plotting flag. The editor also exposes scatter horizontal and primary/secondary vertical axis bounds and major units.

## Verification

The latest `npm run regression:core` passed: TypeScript check, 1,113 tests, a 10,000-formula dependency benchmark, and a 10,000-cell Data Studio benchmark. All 7 desktop code tests passed. The current Windows portable app and Inno Setup installer were rebuilt successfully. No live app, browser UI, Excel, or Calc GUI was used. Re-run this suite after any further edits.

## Release status

The Windows installer, portable archive, source archive, and GitHub-ready repository ZIP are built. SHA-256 checks and ZIP integrity checks passed, and the repository ZIP contains the source, build folder, installer, and handoff. No app was launched. For future source changes, rerun `npm run regression:core`, rebuild Windows artifacts with `desktop/build_windows.ps1 -SkipTests`, then run `python scripts/package_source_release.py` and reverify the checksums. Keep follow-up work code-only unless the user changes the explicit no-live-UI-testing constraint.

## Useful files

- `docs/RELEASE_NOTES_V12_0_209.md`
- `docs/FEATURE_VERIFICATION_V12.md`
- `docs/EXCEL_COMPATIBILITY_AUDIT.md`
- `tests/charts_native_v108.test.mjs`
- `src/engine/charts.ts`, `src/engine/xlsx.ts`, `src/engine/ods.ts`, `src/main.ts`
