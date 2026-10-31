# GridForge 12.0.20

XLSX formula export now preserves Excel's `_xlfn.` compatibility prefix for supported newer functions that were missing from the serializer list: `REGEXEXTRACT`, `GROUPBY`, `PIVOTBY`, `TRANSLATE`, and `IMAGE`. This lets newer Excel versions recognize the functions and gives older builds the expected unsupported-function marker.

This corrects formula serialization metadata only; it does not expand those functions' calculation behavior or provide service-dependent results offline.

Verification: TypeScript static check, production build, and Windows packaging with `-SkipTests`. No tests, app launch, live UI, or Excel GUI checks were run.
