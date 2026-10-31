# GridForge cloud workbook setup

GridForge can browse, open, and explicitly save XLSX workbooks through OneDrive, Google Drive, Dropbox, and Box. Cloud connections are optional; ordinary local workbook use remains available offline. The connector runs in the installed Windows desktop app so refresh tokens and the Box client secret can be encrypted with Windows DPAPI for the current user.

## First-time setup

1. Open **File → Cloud**, select a provider, and register a public desktop OAuth application with that provider.
2. Enable the provider's file API and use the exact loopback redirect required below.
3. Paste the provider's OAuth client ID in GridForge, then choose **Connect account**. Sign-in opens in the system browser.
4. Browse folders and the XLSX/JSON files they contain. Select a folder and choose **Open folder**, or double-click it; use the breadcrumb path to move back up. Save a new copy into the folder currently open. Opening a file only downloads it to the local GridForge session. Use **Save selected workbook** to explicitly update a cloud workbook or **Save a new copy** to create another XLSX file.

The shipped GridForge source contains no provider client IDs or application secrets. You supply your own app credentials; Box's client secret and provider refresh tokens are protected locally with Windows DPAPI. Each provider needs its own app registration and consent configuration.

## Provider registration

### OneDrive

- Register a **mobile and desktop** public client in Microsoft Entra and enable the Microsoft Graph delegated permissions `User.Read` and `Files.ReadWrite`.
- Register a loopback redirect URI ending in `/oauth/callback`. GridForge binds the callback to the local device and uses authorization code with PKCE. Microsoft's localhost redirect matching permits ephemeral ports; see [redirect URI guidance](https://learn.microsoft.com/entra/identity-platform/reply-url) and [desktop authentication flows](https://learn.microsoft.com/en-us/entra/identity-platform/v2-app-types).
- GridForge browses folders and workbooks in OneDrive; new copies are saved in the current folder, and updates use the listed item eTag.

### Google Drive

- Create a **Desktop app** OAuth client, enable Google Drive API, and use the installed-app loopback redirect flow. Google supports PKCE and loopback redirects for desktop apps: [desktop OAuth flow](https://developers.google.com/youtube/data/live/guides/auth/installed-apps).
- GridForge requests the narrower `drive.file` scope rather than access to every Drive file. This limits the list to files GridForge created or that the user opened/shared with it. The scope definition is in [Google Drive OAuth scopes](https://developers.google.com/identity/protocols/oauth2/scopes).
- Some Google OAuth clients require their desktop client secret at sign-in. For an installed application, Google notes that this value cannot be treated as a confidential secret; GridForge keeps it in the local provider settings and does not put it in the repository.
- Before saving over an opened file, GridForge checks the Drive version. If the version changed, it asks the user to reopen the latest copy.

### Dropbox

- Create a Dropbox API app with file content read/write and metadata read scopes. Enable OAuth code flow with PKCE and offline refresh tokens; Dropbox documents this in its [OAuth guide](https://docs.dropboxapi.com/dropbox-api/docs/oauth).
- Register the redirect URI **exactly** as `http://localhost:53682/oauth/callback` in the Dropbox app console and keep the same value in GridForge. If that port is already in use, choose another free localhost port, register it, and enter the matching URI.
- Dropbox saves use the file revision supplied by Dropbox. If another change has replaced that revision, Dropbox rejects the update rather than overwriting the newer file.

### Box

- Create a Box Platform app that uses **OAuth 2.0 user authentication**. Enable read and write access to files and folders, and register the loopback redirect URI `http://localhost:53683/oauth/callback`. Box permits localhost redirect URIs to use a variable port; GridForge currently uses port 53683 for the callback. See Box's [OAuth app setup](https://developer.box.com/guides/authentication/oauth2/oauth2-setup/) and [OAuth 2.0 guide](https://developer.box.com/guides/authentication/oauth2/).
- Enter both the app's client ID and client secret in GridForge. The secret is stored through Windows DPAPI and is not kept in browser local storage. Box workbook requests run in the desktop process, so the app does not require a CORS allow-list entry for GridForge's random local WebView origin; Box's [CORS setup guidance](https://developer.box.com/guides/applications/platform-apps/create/) explains the origin allow-list for browser-side API clients.
- GridForge browses workbook files and folders in Box, downloads selected files, creates copies in the current folder, and updates an opened file only when its current eTag matches the version GridForge read. Box's [folder items](https://developer.box.com/reference/get-folders-id-items/), [file download](https://developer.box.com/reference/get-files-id-content/), [file upload](https://developer.box.com/reference/post-files-content/), and [new file version](https://developer.box.com/reference/post-files-id-content/) endpoints define these operations.
- Box's simple upload endpoint is limited to 50 MB in GridForge. Larger uploads require Box's chunked upload flow and are not yet implemented. The native bridge currently transfers file bytes through the desktop WebView bridge, so large-file transfer performance has not been measured.

## Current limits

- Cloud browsing, account login, and credential protection are implemented for the Windows desktop build. The browser-only development build does not persist cloud tokens. Provider credentials, redirects, API scopes, and network policies must be configured in each account's developer console; code-only checks do not verify an account-specific sign-in.
- GridForge lists XLSX and GridForge JSON files and supports folder navigation across all four providers. It does not yet use Google Picker, import native Google Sheets as editable sheets, sync continuously, resolve simultaneous editing, or provide provider-specific version history controls. Chunked uploads are not implemented.
- Saving to a cloud workbook is an explicit action. There is no background autosave or real-time coauthoring.
- GridForge XLSX export does not include GridForge-only Draw ink; the app asks before omitting it. Save GridForge JSON locally to preserve that artwork.
- Disconnecting deletes the encrypted local token. To revoke provider consent completely, remove GridForge from that provider account's connected-app/security settings.
