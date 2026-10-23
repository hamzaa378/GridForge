# GridForge v6 release evidence

**Release:** 6.0.0  
**Scope:** Add 90 function names, dotted formula grammar, reference-aware evaluation, richer validated v2 action macros, new tests and documentation.

## Test matrix

| Stage | Result | Notes |
|---|---|---|
| `npm run check` | PASS | TypeScript type checking |
| `npm test` | **340/340 PASS** | Existing 241 tests retained and extended; 90 new per-function tests plus dedicated macro tests and dynamic dependency tests |
| `npm run benchmark` | PASS | 10,000 formulas; incremental recalculation statistics generated during run |
| Browser regression, Chromium | **24/24 PASS** | Older spreadsheet interactions plus v6 galleries, dotted functions, VBA-safe import, action macro previews and replay |
| Python prebuilt launcher | PASS | No Node modules or external pip modules required for ordinary use; verify again after extraction |
| Browser storage reload | **NOT VERIFIED IN HOSTED BROWSER** | Hosted browser blocks normal localhost navigation; repeat on Windows |
| Real-world Microsoft Excel VBA | **NOT TESTED/NOT SUPPORTED** | GridForge does not execute Excel VBA |

`npm run regression` is a fail-fast pipeline that runs TypeScript checking, Node unit tests, dependency benchmark and isolated Chromium UI regression. Browser tests use an isolated server/profile. On systems where Playwright cannot find Chromium, install it with `python -m playwright install chromium`.

The 90-function demo is generated from the actual test examples. Python-only launch, XLSX import/export, and v1 macro compatibility are separately covered.

## Known limitations

- Excel VBA, legacy `.xls`, cross-sheet formula references and dynamic array spilling are unsupported.
- Some new formulas support narrower parameter modes or approximate numerical answers compared with Microsoft Excel. Consult the generated function reference and v6 compatibility notes.
- Macro v2 preflights a full workbook copy and commits one transaction per affected sheet; undo across multiple sheets must be performed separately.
- Browser-local persistence and Microsoft Excel round-trip fidelity should be verified on the user's environment.

## Final measured benchmark (local build)

The full regression command completed all four stages successfully. Environment-specific latencies must be taken from the benchmark output rather than copied as general performance guarantees. The benchmark covers 400 rows × 52 columns, 10,000 formula cells, and selective invalidation after source edits.

## Raw reproducible evidence

- [`V6_REGRESSION_OUTPUT.txt`](V6_REGRESSION_OUTPUT.txt): full four-stage runner stdout and individual Node/browser test outcomes.
- [`V6_PYTHON_ONLY_SMOKE.txt`](V6_PYTHON_ONLY_SMOKE.txt): HTTP and static module loading checks using Python's standard library.
- These logs were generated locally; they are not evidence that the application has been deployed on GitHub Pages or tested with Microsoft Excel's VBA runtime.
