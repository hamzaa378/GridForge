# GridForge v10.4 — Development handover

**Base:** v10.3 Excel-style UI + calculation repair release. Core engine and XLSM/VBA code preserved. The release is a **registered-name compatibility and ribbon usability** update, not a full external-data engine.

## Source responsibilities

- `src/engine/functions_connected.ts`: ten explicit connected-function definitions and dependency metadata; safely reject unsupported offline evaluations.
- `src/engine/functions.ts`: combines the existing 506-function implementation registry with ten connected signatures.
- `src/engine/categories.ts`: adds External & Connected category.
- `src/main.ts`: ribbon tabs, command visibility, Ctrl+F1 collapse, arrow navigation, themes and existing spreadsheet.
- `public/styles.css`: ribbon layout and light/dark styles.
- `scripts/excel_connected_windows.py`: opt-in missing-connection test using installed Windows Excel; no service requests.
- `scripts/smoke_ribbon_v10_4.py`: independent 19-check Chromium interface smoke, does not use browser storage reload.
- `tests/connected_v10_4.test.mjs`: registered names, categorized warnings, validation and safe error outcomes.
- `tests/parity_506_coverage.test.mjs`: freezes coverage of prior 506 locally implemented names; excludes connected signatures from that fixture set.

## Build and run

`python run_gridforge.py` serves precompiled `dist/` with stdlib only. Build with Node.js 20+: `npm install`, `npm run build`.

`npm run regression:core` covers TypeScript, engine tests and two 10k benchmarks. `npm run regression:ribbon` additionally requires Chromium/Playwright. `python scripts/test_xlsm_browser.py` and `python scripts/test_vba_bridge_protocol.py` verify VBA preservation/protocol security. Full browser regression remains available but the hosted invocation stalled late after 23 passes.

## Remaining work / limitations

1. Connected functions are recognized, not functional. Future providers need user consent, service configuration, data licenses and security threat review.
2. Excel parity with actual connected services is unverified. Run the Windows comparison only as documented; do not claim it tests successful connected output.
3. Dynamic-array spilling, advanced workbook object-model parity, real Excel VBA execution and some specialized finance/statistics conventions remain incomplete.
4. Browser localStorage persistence after real refresh/port change needs ordinary Windows browser acceptance testing; hosted browser blocks navigation.
5. Full 506-function differential parity was last measured against saved Excel reference at v10.3, not rerun for this release.
6. Preserve a single handover guide at future checkpoints and never silently replace the parser with a formula library.
