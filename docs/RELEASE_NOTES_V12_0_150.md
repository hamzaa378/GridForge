# GridForge v12.0.150 — Locale numeric formula parsing

- Formula entry now normalizes localized digits and decimal commas using the system number locale.
- In decimal-comma locales, valid three-digit grouping marks are removed from numeric literals before parsing. Quoted text is left unchanged, and a separator is removed only when it forms a valid grouped number.
- ODS chart export/import, common chart previews, and the previous ODS formatting and formula improvements are included in this package.
- Limitations remain for locale-specific function names and separators beyond comma/semicolon, locale-specific date parsing, conditional formatting in ODS, and full external chart editing. The formula change is parser-level and does not establish full Calc compatibility.
- Static TypeScript build and Windows installer/source packaging are run for this release. No automated tests, app launch, live UI, or LibreOffice GUI round-trip is run.
