# GridForge v12.0.187 — ODS dynamic-array function interoperability

ODS formula import/export now maps additional implemented dynamic-array functions to Calc’s documented `COM.MICROSOFT` compatibility namespace: `TAKE`, `DROP`, `CHOOSECOLS`, `CHOOSEROWS`, `HSTACK`, `VSTACK`, `TOCOL`, `TOROW`, `WRAPROWS`, `WRAPCOLS`, and `EXPAND`. These functions use GridForge’s existing bounded matrix/spill serialization path. Calc documents these functions beginning with LibreOffice 25.8; `XLOOKUP`, `XMATCH`, `SEQUENCE`, `FILTER`, `SORT`, `SORTBY`, and `UNIQUE` are available from 24.8, while `TEXTSPLIT` is documented from 25.8.

This extends function recognition for newer Calc releases, but does not establish full parity for calculation semantics, dynamic-array spill conflicts, locale, error behavior, or older Calc versions.

Verification for this source release: TypeScript check, production frontend build, Windows portable and installer packaging with tests skipped, and source archive packaging. No automated tests, live UI, Windows runtime, or LibreOffice GUI round-trip were run.
