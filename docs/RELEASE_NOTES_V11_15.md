# GridForge 11.15.0 Release Notes

## Stock chart

- Added OHLC stock candles using headers for Open, High, Low and Close, in any column order, with an optional label/date column.
- Incomplete rows are skipped. Invalid high/low bounds display a useful message instead of failing chart rendering.
- Stock charts are available from the Insert chart gallery, editor and chart-design ribbon.

## Installer update

- Updated GridForge desktop metadata and the Inno Setup configuration to version 11.15.0.
- Compiled successfully as `GridForge-Setup-11.15.0-win64.exe` with Windows test phases skipped. Compilation does not verify installation or launch behavior. SHA-256 is recorded in `docs/DESKTOP_RELEASE_MANIFEST.json`.

## Verification and limits

- Stock parsing/rendering and OOXML fallback have source tests.
- Specialized charts, including stock, still export to XLSX using a clustered-column fallback.
- No live UI or Windows tests are run.
