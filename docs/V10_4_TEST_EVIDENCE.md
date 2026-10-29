# GridForge v10.4 — Verified evidence and test limits

| Test | Result | Evidence |
| --- | --- | --- |
| TypeScript type check / build | PASS | `npm run check`, `npm run build` |
| Local unit tests | **834/834 PASS** | `evidence/V10_4_UNIT_TESTS.log` |
| 10,000-formula recalculation benchmark | PASS: single edit recalculated 26 cells | `evidence/V10_4_CORE_REGRESSION.log` |
| 10,000-cell Data Studio edit, undo and unrelated-value integrity | PASS | `evidence/V10_4_CORE_REGRESSION.log` |
| Isolated fresh-profile ribbon checks | **19/19 PASS** | `evidence/V10_4_RIBBON_TESTS.log` |
| Full hosted browser regression | INCOMPLETE: 23 checks passed, later large-workbook stage stalled | `evidence/V10_4_FULL_BROWSER_PARTIAL.log` |
| XLSM VBA binary preservation | PASS (synthetic, not Windows Excel VBA runtime) | `evidence/V10_4_XLSM_TESTS.log` |
| VBA bridge request security checks | PASS (8) | `evidence/V10_4_VBA_PROTOCOL.log` |
| Python stdlib-only local serving | PASS | `evidence/V10_4_PYTHON_SMOKE.log` |
| Saved 506-function Excel-reference snapshot (local re-evaluation only) | 450/456 comparable; 32 resolved from v10.2, no regressions; historical reference | `evidence/V10_4_PARITY_SNAPSHOT.log` |
| Real Windows Excel comparison of connected names | NOT RUN | `scripts/excel_connected_windows.py` for user's Windows machine |

The original 506 local formula fixtures remain frozen and covered. Ten new registered names are classified as external-only and not counted as locally executable, nor as successful Excel parity. The prior v10.3 Excel comparison (450/456 vs user-generated Excel 16.0 snapshot) is historical, not retested in this version.
