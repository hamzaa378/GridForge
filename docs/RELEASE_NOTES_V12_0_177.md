# v12.0.177 — Calc distribution formula interoperability

ODS formula export now uses Calc's documented `COM.MICROSOFT` names for the implemented chi-square (`CHISQ.*`), F (`F.*`), and Student t (`T.*`) distribution and test functions. This includes left/right-tail forms, inverse functions, and tests. Calc can therefore recognize these formulas when opening ODS files, and GridForge's ODS importer removes the namespace when reopening them.

These are name-compatibility mappings only. Calculation differences in error values, boundary handling, precision, and argument edge cases remain possible. No LibreOffice application round-trip was run.

Calc documents these namespace forms in its [statistical functions part one](https://help.libreoffice.org/latest/en-US/text/scalc/01/04060181.html), [part two](https://help.libreoffice.org/latest/en-US/text/scalc/01/04060182.html), and [part five](https://help.libreoffice.org/latest/en-US/text/scalc/01/04060185.html).
