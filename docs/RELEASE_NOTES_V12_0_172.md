# v12.0.172 — larger ODS chart data caches

Native ODS chart packages now include up to 24 series and up to 6,000 series values per chart, scaling the cached row count to keep each chart bounded. This increases the data window available to Calc and other ODF readers for large charts while retaining limits on total XML growth.

Charts that exceed this cache budget are truncated in other applications; GridForge-authored charts retain their full original range and settings in GridForge metadata. No automated tests or LibreOffice UI round-trip were run.
