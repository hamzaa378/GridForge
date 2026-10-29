# GridForge v11.15 Feature Matrix

| Feature | Implementation | Verification |
|---|---|---|
| Stock OHLC chart | Header-driven Open/High/Low/Close mapping; optional date column; incomplete rows skipped; high/low envelope validation; up/down candles | TypeScript check and focused source tests |
| Installer version | Inno Setup script, desktop app metadata and release manifest are set to 11.15.0 | Compiled successfully with Windows test phases skipped |
| Existing v11.14 tools | Funnel/Treemap charts and Descriptive Statistics report remain included | Full source suite |
| XLSX chart compatibility | Stock and specialized charts use clustered-column fallback | Source serialization tests; no Excel GUI tests |
| Test scope | Source checks only; 975 automated code tests passed; no live UI or Windows tests | User restriction honored |

## Limits

- Stock charts require one header row naming Open, High, Low and Close; an optional leading column supplies labels. Incomplete numeric rows are skipped.
- Native XLSX export does not preserve stock-candle chart semantics yet; export uses a standard clustered-column chart.
- No installer was tested by installation or launch; compilation alone is not an installation/GUI test.
