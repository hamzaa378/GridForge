# GridForge v12.0.200 — Native ODS color-scale thresholds

ODS conditional color scales now retain each Calc threshold type and value in native `calcext` entries: minimum, maximum, fixed value, percent, percentile, and formula. Formula stops are translated to OpenFormula syntax. Import restores the corresponding threshold definitions from native ODS even when GridForge worksheet metadata is removed. Default three-color scales now use a median percentile midpoint, matching Calc's documented default.

The existing v12.0.199 conditional-formatting work also provides XLSX icon-set and custom scale round-trips, configurable threshold controls, per-rule removal, stable new-rule priorities, and cached percentile sorting during rendering. Native ODS icon sets, rule priority/stop-if-true handling, complete blank-cell coverage, and every Calc UI operation remain incomplete. No live UI or LibreOffice GUI testing was run.

Verification: TypeScript check/build, 49 focused ODS regressions, native-only ODS threshold round-trips, the full 1,087-test code suite, and core benchmarks.
