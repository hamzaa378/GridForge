# v12.0.179 — Calc text and criteria formula mappings

ODS export now qualifies GridForge's implemented `MAXIFS`, `MINIFS`, `CONCAT`, and `TEXTJOIN` formulas with Calc's documented `COM.MICROSOFT` namespace. This lets Calc recognize these modern spreadsheet functions in ODS formulas. GridForge's ODS importer strips that namespace on return.

Calc's criteria matching can depend on calculation settings such as wildcard or regular-expression behavior. GridForge continues to calculate its own workbook using its existing matching rules; this export mapping does not make those engines identical. No Calc application round-trip was run.

Calc documents the namespace in its [MAXIFS](https://help.libreoffice.org/latest/en-US/text/scalc/01/func_maxifs.html), [MINIFS](https://help.libreoffice.org/latest/en-US/text/scalc/01/func_minifs.html), [CONCAT](https://help.libreoffice.org/latest/en-US/text/scalc/01/func_concat.html), and [TEXTJOIN](https://help.libreoffice.org/latest/en-US/text/scalc/01/func_textjoin.html) references.
