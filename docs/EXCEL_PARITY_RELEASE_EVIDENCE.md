# GridForge v10.1: Excel Compatibility Audit Evidence

## Test machine and scope

- Host environment: Linux with Node.js v22 and system Chromium.
- Microsoft Excel desktop: **not installed**, so the project does **not** claim actual Excel numerical parity.
- Formula source: custom TypeScript parser/evaluator, 506 registered functions.
- Test fixture count: 140; 136 scalar expected-value checks; 4 known dynamic-array gaps.

## Executed checks

| Check | Result | Evidence |
| --- | --- | --- |
| TypeScript compilation | PASS | `npm run check`, `npm run build` |
| Unit tests and formula reference fixtures | **816/816 PASS** | `docs/EXCEL_PARITY_UNIT_TEST_LOG.txt` |
| Static scalar expected values | **136/136 MATCH** | `docs/PARITY_GRIDFORGE_RESULTS.json` |
| Dynamic-array spill cases | **4 documented gaps** | `tests/excel_parity_cases.json` |
| Chromium UI interactions | **30/30 PASS** | `docs/EXCEL_PARITY_BROWSER_LOG.txt` |
| 10,000-formula dependency benchmark | PASS | `docs/EXCEL_PARITY_SUPPLEMENTAL_LOG.txt` |
| 10,000-cell editing benchmark | PASS | Same supplemental log |
| Python-only compiled-app smoke test | PASS | Same supplemental log |
| XLSM edit and VBA binary preservation | PASS | Same supplemental log |
| Local Excel bridge security checks | PASS | Same supplemental log |
| Actual Microsoft Excel recalculation | **NOT EXECUTED** | Requires Windows with Excel desktop and pywin32 |
| Actual Excel VBA execution | **NOT EXECUTED** | Requires Windows with Excel desktop, user-approved macro |
| Combined regression invocation | **INCOMPLETE / STALLED** | See `docs/EXCEL_PARITY_COMBINED_ATTEMPT.txt` |

The combined test attempt stalled partway through Chromium, so no claim of a complete unified run is made; all named stages passed independently. The browser's real storage-reload test was skipped due to hosted navigation restrictions.

## Corrected defects

1. Rounding of negative halves (`ROUND(-1.5,0)`) now agrees with Excel reference examples.
2. `SUM`, `COUNT` and related aggregations now distinguish directly typed numeric text/booleans from referenced text/booleans.
3. Criteria evaluation now includes `*`, `?` and escaped `~` wildcard patterns for text matching.

## Excel run workflow

On Windows with installed Microsoft Excel:

```powershell
py -3 -m pip install pywin32
py -3 scripts\excel_parity_desktop.py
py -3 scripts\excel_parity_compare.py
```

The test generates `docs/PARITY_EXCEL_RESULTS.json`, `docs/PARITY_COMPARISON.md` and `docs/PARITY_COMPARISON.json`. These outputs are deliberately **not bundled**, because real Excel was not available for this release's tests.

For reporting externally, do not use 136/136 as "Excel agreement". It is agreement with an expected-value fixture set, not direct Excel calculation. Add user workbook fixtures and broaden the differential tests before claiming general equivalence.
