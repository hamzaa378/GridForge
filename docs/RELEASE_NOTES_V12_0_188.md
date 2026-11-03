# GridForge v12.0.188 — ODS RANDARRAY interoperability

ODS formula import/export now maps the implemented `RANDARRAY` function to Calc’s documented `COM.MICROSOFT.RANDARRAY` compatibility name. Bounded array results continue through the existing matrix formula and cached-value path. LibreOffice documents RANDARRAY from version 24.8.

This improves formula recognition by Calc but does not guarantee matching volatile recalculation, random-number distribution, dimensions, errors, locale, or precision. A Calc GUI round-trip remains necessary for comparison.

Verification for this source release: TypeScript check, production frontend build, Windows portable and installer packaging with tests skipped, and source archive packaging. No automated tests, live UI, Windows runtime, or LibreOffice GUI round-trip were run.
