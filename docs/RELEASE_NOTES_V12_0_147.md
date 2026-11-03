# GridForge v12.0.147 — OpenDocument AutoFilters

- ODS export writes basic database-range filters with header detection and filter buttons.
- Contains queries and selected-value criteria across multiple columns are encoded with OpenDocument filter conditions.
- ODS import restores the supported filter range, query column, and selected values.
- Export stops with a clear message for a value criterion with no selected values, which has no safe equivalent in this supported OpenDocument filter subset.

Advanced filter operators, sort state, and more complex database ranges remain unsupported. The build and installer skip tests; no LibreOffice GUI or live UI check was performed.
