# v12.0.180 — Calc array formula mappings

ODS export now uses Calc's documented `COM.MICROSOFT` namespace for GridForge's implemented `SORT`, `UNIQUE`, and `TEXTSPLIT` functions. Calc has native implementations for these names in recent releases and documents their array-formula entry behavior. GridForge strips the namespace on ODS import.

LibreOffice release age matters: the cited help pages say `SORT` and `UNIQUE` are available from 24.8 and `TEXTSPLIT` from 25.8. Older Calc versions may not recalculate them. GridForge and Calc can also differ in array-entry/spill semantics and argument edge cases. No Calc round-trip was run.

Calc references: [SORT](https://help.libreoffice.org/latest/en-US/text/scalc/01/func_sort.html), [UNIQUE](https://help.libreoffice.org/latest/en-US/text/scalc/01/func_unique.html), and [TEXTSPLIT](https://help.libreoffice.org/latest/en-US/text/scalc/01/func_textsplit.html).
