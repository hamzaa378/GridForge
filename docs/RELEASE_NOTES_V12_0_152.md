# GridForge v12.0.152 — ODS source-formula preservation

- When a formula cannot be translated to OpenFormula, ODS export continues to write the calculated cached value and now also stores the original formula in a GridForge-namespaced cell attribute.
- ODS import prefers that preserved source text for these cells, so GridForge can reopen and edit formulas outside its OpenFormula translation subset.
- Other ODF applications continue to see the cached value for those formulas; this metadata does not provide cross-application calculation support.
- Formula attributes are bounded to 8,192 characters. Formulas within the supported translation subset continue to use standard OpenFormula.
- Earlier formula locale, half-point font, border, page-layout, and chart changes remain included. Static build and packaging do not run automated tests or launch the live UI.
