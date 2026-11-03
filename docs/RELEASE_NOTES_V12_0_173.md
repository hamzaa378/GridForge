# v12.0.173 — Calc-compatible Excel function formulas in ODS

ODS formula export now writes the Calc-specific `COM.MICROSOFT` namespace for `IFS`, `SWITCH`, `CEILING.MATH`, and `FLOOR.MATH`. These functions are supported by LibreOffice Calc as Microsoft-compatibility functions, although they are outside OpenFormula 1.3. GridForge continues to reopen them as the original function names.

This improves formula portability to LibreOffice versions that support these functions. Other Excel-only functions outside OpenFormula continue to use cached values with GridForge source formulas in metadata. The mappings were code-checked but not opened in LibreOffice. Calc documents the `COM.MICROSOFT` namespace for these functions: [IFS and SWITCH](https://help.libreoffice.org/latest/en-US/text/scalc/01/func_ifs.html), [CEILING.MATH](https://help.libreoffice.org/latest/en-US/text/scalc/01/func_ceiling.html), [FLOOR.MATH](https://help.libreoffice.org/latest/en-US/text/scalc/01/func_floor.html).
