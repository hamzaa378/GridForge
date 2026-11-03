# GridForge v12.0.196 — Conditional data-bar controls

Conditional data bars now retain automatic/middle/hidden axis placement, axis color, and minimum/maximum bar-length percentages in the rule model and the editor. XLSX and ODS import/export preserve these settings; the sheet preview applies the configured length range. Validation rejects invalid percentages and inverted length intervals.

Full Calc and Excel behavior for axis rendering, rule priority, formula thresholds, borders, icon sets, and advanced data-bar direction remains incomplete. No live UI or LibreOffice GUI test was run.

Verification: TypeScript build and focused ODS/XLSX conditional-format regression tests passed.
