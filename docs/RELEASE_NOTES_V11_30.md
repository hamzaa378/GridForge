# GridForge 11.30.0

## Cloud workbook storage

- Added explicit browse, open, refresh, save-current, and save-copy operations for OneDrive, Google Drive, and Dropbox.
- Sign-in uses provider OAuth authorization code with PKCE in the system browser. Provider client IDs are configured by the user; no provider app secret or access token is embedded in the source or installer.
- On Windows, refresh credentials are protected for the current Windows user with DPAPI.
- Save-over checks use OneDrive eTags, Dropbox revisions, or a Google Drive version preflight to reduce accidental overwrites.
- Ordinary local spreadsheet use remains offline. Cloud access needs an internet connection and user-created provider app registrations.

## Scope and known limits

Cloud file listing is limited to visible XLSX and GridForge JSON files in the provider root. Google Drive uses the limited `drive.file` scope, which does not expose every existing file automatically. Native Google Sheets, folder browsing, continuous sync, simultaneous editing, provider sharing controls, and background autosave are not included. XLSX export still warns before omitting GridForge-only drawing ink.

## Build and verification

- `npm run check` and `npm run build` passed.
- The Windows installer and portable desktop binaries compiled with `desktop/build_windows.ps1 -SkipTests`.
- No automated tests, Windows runtime tests, application launch, live UI test, or Excel GUI test was run.
- First-time cloud use requires registering an OAuth app with each provider. See [Cloud storage setup](CLOUD_STORAGE_SETUP.md).

