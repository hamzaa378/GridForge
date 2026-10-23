# Regression test handover — GridForge v3.0.1

Based on the GitHub-ready v3 source. This update is test-only. Do not confuse this with a release implementing more spreadsheet functions. Existing runtime features and data formats are unchanged.

**Entry points:** `npm run regression` executes `scripts/regression.mjs`, which invokes the TypeScript compiler, 65 Node tests in `tests/*.test.mjs`, and `scripts/benchmark.mjs`, then `scripts/regression_browser.py` using an isolated Chromium profile and randomly selected server port. `npm run regression:core` skips browser dependencies. On Windows `run_regression_windows.ps1` is a simple wrapper.

**New files:** `tests/regression_extra.test.mjs` (8 invariant checks), `scripts/regression.mjs` (cross-platform runner), `scripts/regression_browser.py` (14 UI integration checks), `run_regression_windows.ps1`, `docs/REGRESSION_TESTS.md` (test guide), this handover. Updated `package.json`, `.github/workflows/ci.yml`, `README.md` and release notes.

**Starting from another chat:** Extract the ZIP, open the project root, install Node.js and dependencies, install Python Playwright + Chromium, then run `npm run regression`. For future feature development, add a Node test for engine-level logic and extend `scripts/regression_browser.py` for each UI change. Keep the test suite stable; investigate failures rather than weakening assertions.

**Known limitation:** A managed browser that disallows network navigation uses an inline-module fallback and skips the browser localStorage reload assertion. This is visibly announced; it does not count as a storage-persistence verification. A normal Windows run should use the HTTP mode and verify reload.

**Database:** none. **App data:** browser-localStorage, local JSON export, CSV and XLSX subsets. **Full code state:** same as GridForge v3.0. See `docs/PROJECT_HANDOVER.md` for application modules and architectural limitations.
