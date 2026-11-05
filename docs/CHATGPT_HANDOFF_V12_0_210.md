# GridForge v12.0.210 continuation handoff

## Current change

This working copy advances v12.0.209 to v12.0.210. `TEXTBEFORE` and `TEXTAFTER` match Excel's documented empty-input result, omitted `instance_num` and `if_not_found` defaults, explicit empty-text fallbacks, `match_end`, empty-delimiter behavior, and `#VALUE!` when `ABS(instance_num)` exceeds the text length. `TEXTSPLIT` rejects explicitly empty delimiters, preserves omitted-column row splitting, and pads uneven results with `#N/A` when `pad_with` is omitted, including an explicit omitted argument slot. `FILTER` returns `#CALC!` for no matches when `if_empty` is omitted, while explicit empty text remains an empty result. ODS and XLSX formula round-trips cover these optional slots; ODS maps the text functions through Calc's `COM.MICROSOFT` extension. `NETWORKDAYS.INTL`/`NETWORKDAYSINTL` and `WORKDAY.INTL` use the standard weekend when the optional weekend slot is omitted before holidays; numeric weekend code `0` returns `#NUM!`, an all-weekend `WORKDAY.INTL` mask returns `#VALUE!`, and `NETWORKDAYS.INTL` returns 0 for that mask. `SORT`/`SORTBY`, `SEQUENCE`, `RANDARRAY`, `EXPAND`, `WRAPROWS`, and `WRAPCOLS` honor omitted arguments and documented #N/A padding; `RANDARRAY` rejects equal minimum/maximum bounds. Regression coverage is in `tests/functions_v4.test.mjs`, `tests/functions_v6.test.mjs`, `tests/functions_v10.test.mjs`, and `tests/ods_chart_formula_regression.test.mjs`; implementation is in `src/engine/functions_extra.ts`, `src/engine/functions_v6.ts`, `src/engine/functions_v10.ts`, and `src/engine/ods.ts`.

The prior v12.0.209 scatter chart work remains included: up to 24 independent X/Y pairs, hidden-pair preview/ODS handling, Excel visible-cell plotting for XLSX, per-series style preservation, scatter axis settings, and numeric point-ID header inference.

## Verification and release

All 1,123 source tests, 7 desktop code tests, the TypeScript build, both 10,000-item benchmarks, and Windows portable/installer compilation passed. Generate the portable and GitHub source archives with `python scripts/package_source_release.py`. No live app/UI or Excel/Calc GUI testing was run.

## Useful files

- `docs/RELEASE_NOTES_V12_0_210.md`
- `docs/EXCEL_COMPATIBILITY_AUDIT.md`
- `tests/functions_v4.test.mjs`
- `src/engine/functions_extra.ts`
