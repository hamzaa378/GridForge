# GridForge v10.0.0 — release verification

## Tested results

| Check | Result | Evidence |
| --- | --- | --- |
| Full release regression | **7/7 stages PASS** | `docs/V10_FULL_REGRESSION_OUTPUT.txt` |
| TypeScript strict compilation | PASS | `npm run check` |
| Automated formula, workbook, XLSX/XLSM and feature suite | **675/675 PASS** | `npm test`, release log |
| 100 new function registration and callability | **100/100** | `tests/functions_v10.test.mjs` |
| Real Chromium browser functional suite | **30/30 PASS** | `scripts/regression_browser.py`, release log |
| Excel-like theme and workspace controls | PASS | New v10 browser checks |
| 10,000-formula dependency benchmark | PASS | `npm run benchmark` |
| 10,000-cell Data Studio preview/apply/undo | PASS | `npm run benchmark:tools` |
| XLSM edited-cell/formula round-trip and VBA preservation | PASS | `python scripts/test_xlsm_browser.py` |
| VBA bridge authentication and request security | PASS | `python scripts/test_vba_bridge_protocol.py` |
| Python-only compiled release startup | PASS | `python scripts/python_smoke.py` |

## Explicit limitations and unverified areas

- The hosted Chromium environment blocks actual navigation to local origins; **reload-based persistence of theme, workspace and workbook was not independently verified**. The code uses versioned `localStorage` keys and head-time initialization; please verify on Windows.
- **Desktop Excel VBA execution is NOT verified**. Excel desktop and pywin32 are required. Bridge request validation was tested only; do not claim VBA production compatibility on that basis.
- The 506 count is registered names, not 506 feature-identical Excel implementations. Dynamic spilling, certain financial/date conventions, Lambda closures, and external data services differ.
- Light and dark themes are application-level palettes; unusual user-authored cell colors can have low contrast and are intentionally not overwritten by the theme.
- Extensive numeric cross-validation against desktop Excel, Edge/Safari and assistive technologies remains future work. Passing tests do not guarantee absence of bugs.

## Running all checks

```powershell
npm install
npm run check
npm test
npm run benchmark
npm run benchmark:tools
python -m pip install playwright
python -m playwright install chromium
npm run regression
python scripts/python_smoke.py
```

CI is prepared for GitHub; deployment has not been independently verified. For a Windows acceptance, also test theme reload, storage reload, Excel XLSM editing and real trusted VBA execution under Excel Trust Center rules.
