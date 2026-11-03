# v12.0.171 — ODS chart axis number formats

ODS chart export now places supported number formats on the embedded chart data and links numeric axes to those source formats. Native ODF import reads linked numeric, currency, percentage, integer, and supported custom formats for x, primary y, and secondary y axes.

Custom formats are limited to the numeric patterns already supported by GridForge's ODS style converter. Date-based chart axes, locale-specific formats, and live source-data linkage remain partial. This was code-checked only; no automated tests or LibreOffice UI round-trip were run.
