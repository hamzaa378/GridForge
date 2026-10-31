# GridForge 12.0.48

Treemap and Sunburst now group categories by the full JSON-encoded hierarchy tuple. Labels containing the visible ` › ` path separator can no longer merge two distinct branches into one category. The local SVG preview carries the exact hierarchy alongside its display labels. XLSX ChartEx export continues to use the original worksheet hierarchy columns.

Verification is limited to static TypeScript checking, Python syntax compilation, and production packaging with Windows tests skipped. No tests, app launch, live UI, Windows runtime, or Excel GUI checks were run.
