# v12.0.178 — Calc probability distribution compatibility

ODS formula export now qualifies the implemented `BETA.DIST`, `BETA.INV`, `GAMMA.DIST`, `GAMMA.INV`, `HYPGEOM.DIST`, and `GAMMALN.PRECISE` functions with Calc's documented `COM.MICROSOFT` namespace. Calc can identify these Excel-compatible functions in ODS formulas, and GridForge strips the namespace when importing them again.

These mappings improve formula recognition but do not promise identical results for every boundary case, error, or precision level. No LibreOffice application round-trip was run.

Calc documents these names in the [statistical functions part one](https://help.libreoffice.org/latest/en-US/text/scalc/01/04060181.html) and [part two](https://help.libreoffice.org/latest/en-US/text/scalc/01/04060182.html).
