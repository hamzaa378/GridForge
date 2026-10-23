# GridForge v8.0 — Reproducible release evidence

## Scope

This release contains 100 new **Excel-style Data Studio actions** (10 categories of 10), retains **306 custom spreadsheet functions**, preserves XLSX/XLSM support and the optional Windows Excel VBA bridge. User-requested bug audit and regression coverage are provided by the new tests and pre-existing tests.

## Results in this Linux build environment

| Stage | Verification | Outcome |
|---|---|---|
| 1 | Strict TypeScript typechecking (`npm run check`) | PASS |
| 2 | Unit tests (`npm test`): original engine, workbook, XLSX, macro suites plus v8 actions | **447 passed; 0 failed** |
| 3 | 10,000-formula selective-recalculation benchmark (`npm run benchmark`) | PASS |
| 4 | 10,000-cell Data Studio preview, apply, undo, unrelated-cell check and oversize rejection | PASS |
| 5 | Isolated Chromium UI regression | **26 passed**, no JavaScript page errors |
| 6 | XLSM preservation and edited-file readback | PASS |
| 7 | Local Excel VBA bridge authentication, origin, input and file preflight security tests | PASS |
| Python-only smoke | Standard-library local server, static JS/CSS MIME and missing-file response | PASS |
| ZIP integrity | Archive listing, CRC check, clean-extraction Python launch | PASS (see release archive verification) |

The final complete run recorded in `docs/V8_FULL_REGRESSION_OUTPUT.txt` reported **`REGRESSION PASS: 7/7 stages in 35.9s`**. An earlier repeated run stalled in the XLSM browser test; the stage also passed when run separately. The final successful run is the release evidence. The test scripts and CI workflow remain in the repository for reproducibility.

The 10,000-cell tool benchmark on this machine reported approximately 45 ms for preview and 53 ms for applying the edits, and verified Undo and unaffected cells. These numbers are environment-specific examples, not a promised latency bound.

## Explicitly unverified on this machine

- Windows Microsoft Excel COM/VBA execution: requires licensed desktop Excel, a trusted macro-enabled workbook and manual confirmation. We only checked the local bridge request validation here.
- Real browser `localStorage` reload from a regular HTTP origin: hosted Chromium restrictions forced inline document mode and skipped this check. Test on Windows.
- Every possible Excel OOXML feature/complex real-world macro workbook: the importer and exporter implement a documented subset, not Excel's entire format.
- Perfect Excel semantics for all possible combinations of 306 registered functions: registration and targeted tests are comprehensive for covered cases, not mathematical proof of full Excel equivalence.

## Commands to reproduce

```powershell
npm install
npm run check
npm test
npm run benchmark
npm run benchmark:tools
npm run regression
python scripts/python_smoke.py
```

For the browser stages, install Python Playwright and Chromium as described in the README. The module-specific test `tests/data_studio_v8.test.mjs` includes a dedicated behavior case for all 100 actions, plus invalid input, formula protection and range-boundary cases. `scripts/regression_browser.py` tests the UI controls and file workflows. `scripts/test_xlsm_browser.py` tests macro-binary preservation, and `scripts/test_vba_bridge_protocol.py` audits local bridge request guards.
