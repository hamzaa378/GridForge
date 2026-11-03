# GridForge v12.0.195 — ODS Calc conditional-format interoperability

ODS now writes color scales and data bars as native LibreOffice Calc `calcext` conditional-format objects, with absolute sheet-qualified target ranges. The importer restores supported two- and three-color scales, automatic or numeric bounds, positive and negative bar colors, solid or gradient bars, and value visibility. A regression removes GridForge-only metadata before import to verify the native ODS representation.

Conditional-format rule precedence, formula-based thresholds, icon sets, data-bar axes and lengths, and styling across blank cells are still incomplete. The ODS extension is Calc-specific and is not part of the ODF core namespace. No LibreOffice GUI round-trip was performed.

Verification: production build and the focused ODS chart/formula regression file passed. No live UI or LibreOffice GUI test was run.
