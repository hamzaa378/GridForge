# GridForge v5.0.0 — Release Verification Evidence

This report distinguishes the tested implementation from unverified Excel/VBA compatibility. It is a reproducible project acceptance record, not a claim of defect-free or certified software.

## Scope of this release

- 216 registered spreadsheet functions: 56 original, 80 added in v4, 80 new in v5.
- Custom TypeScript parser and evaluator, with dependency tracking unchanged.
- Excel-like grouped function browser, search, Recently added filter.
- Secure GridForge action-macro recording and replay for supported cell/style edits.
- `.xlsm` import for supported worksheets, byte-identical original download, **no VBA execution**, edited `.xlsx` export without macros.
- Updated source, compiled browser bundle, zero-dependency Python launcher, examples and GitHub workflows.

## Automated results from this build environment

| Stage | Observed result |
| --- | --- |
| Strict TypeScript typecheck and build | PASS |
| Node unit/engine/document tests | **241 passed, 0 failed** |
| Chromium interactive tests | **21 passed** |
| Test stages (`npm run regression`) | **4/4 PASS** |
| Python launcher and static-site smoke | PASS, standard library only |
| 10,000-formula benchmark | PASS, including selective graph invalidation |
| ZIP integrity and fresh extraction | PASS after fresh extraction (Python smoke and archive CRC verified) |

Browser storage persistence reload was **skipped** because the hosted environment prevented navigation to localhost. A full browser acceptance run on the user's Windows system is required. Previous user-provided Windows logs verified 65/65 core tests and 14/14 browser tests for v3, not v5.

## Benchmark recorded in this environment

- Worksheet: 400 rows × 52 columns = 20,800 addressable cells.
- Populated cells: 10,400; formulas: 10,000.
- Initial graph construction/evaluation: approximately **149.069 ms** on this run.
- Editing one source cell: **26 affected and recalculated**; approximately **0.073 ms** on this run.
- Isolated edit: **1 affected/recalculated**; approximately **0.006 ms**.
- Timing varies across machines; no real-time or production SLA is asserted.

## New v5-specific tests

- Every one of the 80 new functions has a working-formula assertion; category membership is exhaustive and unique.
- Financial/date/engineering invalid-input cases emit typed spreadsheet errors, not unhandled exceptions.
- `ISERR`, `ISERROR`, `ISNA` and `CHOOSE` evaluate branches as intended.
- Secure action-macro JSON rejects arbitrary operations, invalid data, excess steps and overflow.
- Browser records cell and style changes, saves macro, replays it, and exports JSON.
- Browser imports the v5 `.xlsx` formula gallery and calculates representative cells.
- Browser imports a **synthetic** macro-enabled ZIP fixture and verifies that the original file can be downloaded **byte-for-byte**, while an edited `.xlsx` omits the VBA payload.
- The synthetic VBA fixture is NOT a valid executable Excel macro project and therefore does **not** establish full compatibility with real-world VBA projects. Real `.xlsm` files should be separately tested in Excel with macros disabled.

## Reproduction commands

Python-only smoke:

```powershell
python run_gridforge.py --check
python scripts/python_smoke.py
```

Source testing:

```powershell
npm install
npm run regression:core
py -3 -m pip install playwright
py -3 -m playwright install chromium
npm run regression
```

GitHub Pages publication and a full real `.xlsm` interoperability test remain user-side checks. Never enable macros from untrusted spreadsheets.
