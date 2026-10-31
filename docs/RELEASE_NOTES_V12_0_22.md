# GridForge 12.0.22

XLSX formula export now adds `_xlfn.` for `FIELDVALUE` and `STOCKHISTORY`, complementing the newer-function prefixes added in v12.0.20. This preserves Excel's version marker for linked-data and subscription-backed formulas when a workbook is opened in an older Excel build.

GridForge's `FIELDVALUE` remains a limited text-field adapter rather than a linked Excel data-type engine, and `STOCKHISTORY` still requires an external licensed data service to calculate.

Verification: TypeScript static check and Windows packaging with `-SkipTests`. No tests, app launch, live UI, or Excel GUI checks were run.
