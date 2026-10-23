# GridForge v3.0.2 final-release verification

Date: 19 October 2026. The release is based on the fully tested v3.0.1 spreadsheet code; v3.0.2 adds **Python-only launch and submission/documentation improvements**, not new formula-engine behavior.

## Evidence from the user's actual Windows run

| Stage | Result |
| --- | --- |
| `npm run check` (TypeScript check) | PASS |
| `npm test` | PASS, **65 tests, 0 failed** |
| `npm run benchmark` | PASS; 10,000 formulas; one source edit affected/recalculated 26; isolated edit 1 |
| `scripts/regression_browser.py` via `npm run regression` | PASS, **14 checks** in **6.7 s**; real HTTP server, not offline fallback |
| Combined `npm run regression` | **PASS 4/4 stages**, reported **14.5 s** |
| `npm run regression:core` | **PASS 3/3 stages**, reported **4.6 s** |

User's environment: Windows PowerShell, Node v24.21.0. Times will differ across machines and are not SLAs. The Node `[DEP0190]` deprecation warning did not cause a failure.

## v3.0.2 Python-only checks

Run from root with only Python 3:

```sh
python run_gridforge.py --check
python scripts/python_smoke.py
```

Expected: precompiled assets verified, real loopback server serves HTML/JS/CSS and rejects missing files; no pip/npm install required.

## Build and CI commands for a new developer

```sh
npm install
npm run check
npm test
npm run benchmark
npm run regression:core
# For full UI tests, additionally install Python Playwright and its Chromium browser:
npm run regression
```

CI workflow runs Node + Chromium tests on main and PRs. Reproduce on the target machine; the user's Windows results are stronger evidence for that environment than a claim that the app works on all platforms.

## Independent verification of packaged v3.0.2 (this release)

- Python-only launcher `--check`: PASS.
- Python standard-library smoke test: PASS for HTML, CSS, JS modules, MIME types and HTTP 404 handling.
- Local TypeScript strict check: PASS.
- Local Node unit tests: **65/65 passed**, no skipped tests.
- Local 10,000-formula benchmark: PASS, source edit **26** affected and **26** recalculated cells; isolated edit **1**.
- Local full `npm run regression`: **4/4 stages passed**, including **14 Chromium UI checks**. **Hosted Chromium blocked localhost navigation** and used the fallback embedded-page mode, so *browser localStorage reload was explicitly skipped in this environment*. The real HTTP storage test **passed on the user's Windows v3.0.1 run**.
- TypeScript spreadsheet sources did not change between v3.0.1 and v3.0.2; new behavior is a Python static server and documentation.

## Submission caveat

The full user regression was run on v3.0.1. v3.0.2 changes static serving scripts and documentation, so the Python-only server smoke check covers the changed path. Existing engine/UI source and `dist/` are kept intact. Do not claim v3.0.2 full Windows Chromium rerun until it is actually run after extraction.
