# Authenticated HTTPS data import

The Windows desktop app includes **Data → Authenticated Web Data** for a one-time import from an HTTPS CSV or JSON endpoint. It supports no authentication, HTTP Basic username/password, and Bearer tokens. JSON responses can use a JSON Pointer, and the bounded Power Query M interpreter can transform the returned table before it is loaded.

The desktop bridge sends credentials in the request's `Authorization` header. It rejects redirects so credentials are not forwarded to another URL, uses a 15-second timeout, rejects non-HTTPS URLs and embedded URL credentials, and limits the UTF-8 response to 3 MB. The credential exists only in the request/dialog memory and is cleared when the request completes; GridForge does not save it to workbook data or its credential store.

This is a one-time import, not a saved Power Query connection. It does not support OAuth, API-key header customization, pagination, token refresh, database drivers, or query folding. Only send credentials to an HTTPS host you trust. Authenticated URLs are fetched by the Windows desktop bridge; the browser build has no equivalent credentialed connector.
