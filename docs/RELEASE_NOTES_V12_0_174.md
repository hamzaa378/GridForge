# v12.0.174 — Calc statistical formula names

ODS export now emits Calc's documented `COM.MICROSOFT` names for STDEV.S, STDEV.P, VAR.S, VAR.P, and RANK.EQ in addition to the previous IFS, SWITCH, CEILING.MATH, and FLOOR.MATH mappings. These Microsoft-compatibility names are not part of OpenFormula 1.3, but Calc supports them directly. GridForge imports them back as their Excel function names.

Other non-standard formulas still fall back to cached values for external ODF readers. These changes were code-checked but not opened in LibreOffice. See [Calc statistical function names](https://help.libreoffice.org/latest/en-US/text/scalc/01/04060185.html).
