# Local SQLite read-only connector

On Windows desktop, **Data → SQLite Query** opens a native file picker. The selected file is referenced in the desktop process by a random session handle; its path and any SQL text are not stored in the workbook. A saved query/connection manager is not included yet.

The connector accepts one `SELECT` or `WITH` statement, runs it with SQLite `mode=ro`, enables query-only mode, disables extension loading, and uses SQLite's authorizer to deny writes, DDL, attachments, transactions and PRAGMA operations. It also applies an 8-second/5-million-instruction execution budget and caps results at 1,000 rows, 104 columns, 12,000 total cells, and 32,767 characters per text value. BLOB results are rejected. Query results that begin with `=` are loaded as text to prevent a database value from becoming an unintended worksheet formula. Existing worksheet cells are confirmed before they are replaced.

This is a local SQLite reader, not a Power Query connector. It does not support database passwords, network database servers, ODBC/JDBC, SQL Server, PostgreSQL, MySQL, scheduled refresh, credential storage, or OLAP. The session handle expires when GridForge closes. The browser build has no direct local SQLite access.
