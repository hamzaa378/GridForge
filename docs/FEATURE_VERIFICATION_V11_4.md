# GridForge v11.4 Feature Matrix and Verification

This release improves response during chart-heavy worksheet redraws and completes direct axis-title editing in the Format ribbon. It continues the standalone Excel-style grid and ribbon introduced in earlier releases.

| Area | v11.4 behavior | Verification |
|---|---|---|
| Worksheet/chart redraw | Grid paints immediately; visible chart SVGs are queued individually, and stale render queues are ignored | Browser workflow checks chart creation, update and XLSX reopen |
| Chart axes | Horizontal and vertical axis titles can be edited from the Format ribbon; chart dialog options remain available | Chromium browser workflow |
| Desktop server | Loopback-only server selects an ephemeral port for each launch | Desktop unit tests assert loopback binding and a nonzero assigned port |
| Formulas, tables, XLSX/XLSM | Existing supported behavior is retained | See regression output and previous feature matrix |

## Compatibility limits

GridForge is a standalone spreadsheet application. It does not embed, automate or require Microsoft Excel. The feature set is not full Excel parity. Native XLSX chart import/export is limited to the OOXML subset documented in the earlier feature matrix.

## Verification results

Verification completed on Windows 11:

- `npm run regression` — all 8/8 stages passed, including 915/915 Node tests, dependency and Data Studio benchmarks, 32 Chromium checks, the focused table/chart/resize/data-bar workflow, XLSM round-trip and 8 bridge-security checks.
- `python -m unittest discover -s tests -p "test_desktop*.py"` — 7/7 desktop tests passed, including the API/window cycle safeguard and ephemeral-port binding.
- The packaged 11.4 app stayed responsive, painted the worksheet and accepted cell selection. The 11.4 installer compiled and installed to an isolated directory.
- A fresh source-package extraction and install verification are recorded in the release artifacts.

Tests show supported behavior only; they do not establish full Excel compatibility or guarantee that every external WebView2 installation is healthy.
