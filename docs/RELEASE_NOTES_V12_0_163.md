# GridForge v12.0.163 — ODS named number-format colors

ODS import/export now preserves Excel's eight supported named number-format colors (Red, Blue, Green, Yellow, White, Black, Cyan, and Magenta) for the supported numeric custom-format patterns. Colors apply to the positive section and to mapped negative/zero sections, including literal-only zero sections.

The ODF mapping uses number-style text properties. Unsupported `[ColorN]`, conditional format sections, and other advanced number-format syntax remain unconverted.

Verification: TypeScript check, production frontend build, Python packaging checks, portable build, and Windows installer compilation. Automated tests, Windows launch, live UI, and LibreOffice GUI round-tripping were not run.
