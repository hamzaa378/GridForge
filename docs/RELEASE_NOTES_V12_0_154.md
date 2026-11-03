# GridForge v12.0.154 — Native ODS stock charts

- GridForge OHLC/HLC stock charts now export as native ODF `chart:stock` chart objects with category and series data, alongside their SVG previews and preserved GridForge chart settings. Worksheet-source query definitions are also preserved in GridForge metadata; external query URLs are excluded.
- External ODF stock chart objects with a supported contiguous source range can import as editable GridForge stock charts.
- Other native ODS chart classes and GridForge-only worksheet-object round-trips remain included.
- Specialized charts without a standard ODF class still export as previews and GridForge metadata. Full chart wizard/style compatibility, linked-data refresh, and Calc GUI comparisons remain open.
- Static TypeScript build and Windows packaging are run without automated or live UI tests.
