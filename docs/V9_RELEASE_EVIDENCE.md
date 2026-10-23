# GridForge v9 — Release verification and evidence

**Release:** v9.0.0. **Source:** `/src` builds the static `/dist` website. **Test platform:** Linux, Node.js v22.16.0, Python 3.13, Chromium available as system executable. **Regression command:** `GRIDFORGE_CHROMIUM=/usr/bin/chromium npm run regression` (Windows: install Playwright/Chromium and use `npm run regression`).

| Verification | Observed result |
| --- | --- |
| TypeScript strict type checking | PASS |
| Unit/document tests including 100 new v9 formula examples | **550 passed; 0 failed** |
| Function registry | **406 names** (306 v8 + 100 v9), every v9 function directly tested |
| Formula benchmark | 10,000 formulas; 20,800-cell canvas; one source edit **26 cells affected and recalculated**, 0.087 ms in the recorded run |
| Data Studio stress test | 10,000-cell preview ~41.87 ms, apply ~55.89 ms, Undo and unrelated cells verified |
| Chromium UI interaction checks | **28 passed**, including v9 ribbon, search, guide, and formula gallery |
| XLSM binary preservation and updated-cell round-trip | PASS; 3 changed cells and 8 preserved parts in the synthetic fixture |
| Local Excel bridge protocol security | PASS; eight authentication/origin/validation checks |
| **Full regression** | **7/7 stages PASS** in an earlier recorded run; final changes separately passed 4/4 core, 28/28 browser, XLSM and bridge checks |
| Built-in function gap report | Frozen comparison catalog 475 names; 371 covered, 104 missing |

Read `V9_EXCEL_FUNCTION_GAP.md` for methodology and names. Rebuild the report with `npm run functions:gaps`.

## Tested bugs fixed during v9

- `TRUE()` and `FALSE()` previously failed in formulas because boolean literals took precedence over function syntax; the parser now supports both syntax forms.
- Proposed function count included a duplicate `QUARTILE`; replaced it with a real `BINOM.INV` implementation to yield 100 genuinely new registered names.
- Signed engineering-base output now rejects values outside the specified Excel-style signed range (`DEC2BIN(512)` must be a numeric error).
- Invalid SUBTOTAL operation codes no longer silently resolve via modulo 100.
- Switching to Formulas now selects the Functions sidebar tab consistently, and the footer reports the v9 version.

## Limitations of the evidence

- The hosted Chromium browser prevents normal HTTP origin navigation, so the browser regression ran in **offline embedded document mode**. The actual browser `localStorage` reload test was skipped. Repeat that verification in a normal Windows/Chrome browser.
- Actual Microsoft Excel VBA execution on Windows is **not tested** here. The optional Excel bridge can only be fully tested on a trusted Windows machine with desktop Excel installed.
- Newly added statistical/financial compatibility functions have numerical approximations and deliberately bounded argument conventions. Registration and individual tests establish specified examples, not full Excel mathematical equivalence.
- Dynamic-array functions do not spill into cells. Wrap array-valued results with INDEX to return one scalar.
- Comparing against a 475-name snapshot does not mean the complete current Microsoft Excel function catalog has exactly 475 functions.

## Release files

- `examples/v9_100_function_gallery.xlsx` and `.json`: all new formulas evaluated before packaging
- `docs/preview_v9_excel_interface.png`, `preview_v9_onboarding.png`: UI screenshots captured from Chromium
- `docs/V9_EXCEL_FUNCTION_GAP.md` and `V9_GAP_SUMMARY.json`: reproducible snapshot comparison
- `docs/V9_CODE_AUDIT.md`, `V9_UI_QUICK_START.md`, `FUNCTION_REFERENCE.md`: code and learning documentation
- `docs/V9_FULL_REGRESSION_OUTPUT.txt`: full captured regression evidence

**Final verification note:** A repeat combined invocation later stalled in the hosted Chromium session and hit the execution timeout. The independent post-update suite passed (4/4 core, 28/28 browser, separate XLSM and bridge checks); see `V9_STAGE_BY_STAGE_REGRESSION.txt`. The prior full 7/7 result is retained in `V9_FULL_REGRESSION_OUTPUT.txt`.
