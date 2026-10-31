# GridForge 12.0.10

- Adds Box cloud workbook browse, download, create-copy, and update flows in the Windows desktop app.
- Routes Box file operations through the desktop bridge because Box requires browser API origins to be allow-listed and GridForge's WebView listens on a random loopback port. The native bridge accepts only the Box workbook endpoints needed by GridForge and restricts content downloads to Box's download host.
- Protects the Box OAuth client secret using Windows DPAPI and uses the file eTag in `If-Match` when updating an opened workbook.
- Box root-level file listing and simple uploads are supported. GridForge limits Box uploads to 50 MB; Box chunked uploads, folder navigation, and provider-account sign-in/API verification remain outstanding.
- `npm run check`, production build, Python source compilation, installer packaging with `-SkipTests`, and archive integrity checks are the intended code-only verification. No Windows tests, app launch, live UI, or Box account sign-in are performed.
