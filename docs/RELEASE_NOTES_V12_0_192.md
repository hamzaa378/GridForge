# GridForge v12.0.192 — ODS chart and formula compatibility

ODS native charts now export and import OpenDocument settings for missing-cell treatment, hidden-cell inclusion, and series orientation. Supported custom chart data-label number formats use native OpenDocument number styles.

ODS formula import now removes Calc's `COM.MICROSOFT.` namespace only outside quoted strings, preserving literal text that contains the same characters. Implemented `FORECAST.LINEAR` and `Z.TEST` formulas now export with Calc's documented compatibility namespace so Calc can recognize them.

These changes improve file interoperability but do not establish calculation equivalence or full chart fidelity. No LibreOffice Calc GUI round-trip was run.

Verification: TypeScript check passed. Production build, Windows packaging, and source archive packaging are pending. No automated tests or live UI checks were run.
