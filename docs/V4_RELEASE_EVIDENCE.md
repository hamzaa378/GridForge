# GridForge v4.0.0 — Release and regression evidence

Date: 20 October 2026. This document describes the code in this ZIP and distinguishes independent user testing of earlier versions from checks conducted here.

## Implementation verification

| Check | Evidence |
| --- | --- |
| Custom function registry | 136 entries, including exactly 80 additional v4 entries |
| Compilation | `npm run check`, `npm run build` |
| Automated tests | `npm test`: 156 tests, all passing |
| Function-specific tests | 80 independently named formula result checks plus edge/boundary tests |
| Browser smoke | `GRIDFORGE_CHROMIUM=/usr/bin/chromium python scripts/regression_browser.py`: 15 checks passed; storage reload skipped by host navigation restrictions |
| XLSX extension markers | Unit test inspects XML output containing `_xlfn.XLOOKUP` and `_xlfn.IFS`; test confirms quoted text is preserved |
| Dependency graph | Old and new tests cover selective dependency invalidation after formula and cell edits |
| Python-only app | Included precompiled `dist/`; `python run_gridforge.py` requires no third-party Python modules |

**Limit:** The browser smoke run in this environment uses an embedded offline document if HTTP navigation is blocked. As a result, it does not certify actual browser-localStorage reload. Earlier Windows testing passed browser persistence for the v3 baseline; repeat on Windows for v4 release confidence.

- A generated function-gallery workbook was evaluated for all 80 formulas before exporting both JSON and XLSX examples.
- The full four-stage regression also passed in this environment, including Python-only smoke checks. An additional browser scenario imported the bundled 80-function Excel XLSX gallery and verified `LOG10` and `_xlfn.XLOOKUP` evaluation.

## What remains unverified

- Actual opening of exported v4 formula workbooks in desktop Microsoft Excel was not tested.
- Runtime and latency measurements depend on hardware; they are not universal claims.
- Dynamic arrays and full Excel syntax are explicitly out of scope.

## Reviewer verification commands

```powershell
npm install
npm run check
npm test
npm run benchmark
npm run regression
```

First-time full browser setup: `py -3 -m pip install playwright` followed by `py -3 -m playwright install chromium`.

For reviewers who only have Python 3, run `python run_gridforge.py` to use the precompiled app. Tests for the source require Node.js.
