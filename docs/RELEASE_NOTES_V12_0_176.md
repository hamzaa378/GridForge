# v12.0.176 — Calc probability formula interoperability

ODS export now writes Calc's documented `COM.MICROSOFT` namespace for the implemented `BINOM.DIST`, `BINOM.INV`, `CONFIDENCE.NORM`, `CONFIDENCE.T`, `NORM.DIST`, `NORM.INV`, `NORM.S.DIST`, `NORM.S.INV`, `LOGNORM.DIST`, `LOGNORM.INV`, `POISSON.DIST`, `EXPON.DIST`, and `WEIBULL.DIST` functions. This lets Calc recognize these formulas when opening ODS files. ODS import removes that namespace and restores GridForge's regular function names.

The mapping is about Calc recognizing the function names; GridForge's calculation behavior is still not guaranteed to match Calc for all invalid inputs, error codes, precision, locale, and boundary cases. No application round-trip was run.

Calc documents the compatibility names in its [statistical functions reference, parts one](https://help.libreoffice.org/latest/en-US/text/scalc/01/04060181.html), [three](https://help.libreoffice.org/latest/en-US/text/scalc/01/04060183.html), [four](https://help.libreoffice.org/latest/en-US/text/scalc/01/04060184.html), and [five](https://help.libreoffice.org/latest/en-US/text/scalc/01/04060185.html).
