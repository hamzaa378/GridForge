# v12.0.175 — Calc statistical formula interoperability

ODS formula export now writes LibreOffice Calc's documented `COM.MICROSOFT` function names for `MODE.SNGL`, `RANK.AVG`, `PERCENTILE.INC`, `PERCENTILE.EXC`, `PERCENTRANK.INC`, `PERCENTRANK.EXC`, `QUARTILE.INC`, and `QUARTILE.EXC`. Calc can recognize these Excel-compatible functions when opening a GridForge ODS workbook. Import strips the vendor namespace and restores the unqualified Excel function name.

The mapping applies only to GridForge's registered implementations. It does not establish complete behavioral equivalence: error codes, precision, locale, and edge cases can still differ. Unsupported functions continue to export cached values and GridForge's original formula metadata. No Calc application round-trip or GUI check was run.

Calc documents these compatibility names in its [statistical function reference, part four](https://help.libreoffice.org/latest/en-US/text/scalc/01/04060184.html) and [part five](https://help.libreoffice.org/latest/en-US/text/scalc/01/04060185.html).
