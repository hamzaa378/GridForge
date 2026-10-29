# GridForge 11.19 Windows Desktop

- `gridforge_desktop.py`: local-only server and pywebview EdgeChromium desktop wrapper; native Save As API
- `gridforge.ico`: icon for the installed app
- `requirements-windows.txt`: pinned Windows build dependencies
- `build_windows.ps1`: build tested source, standalone EXEs and installer on Windows x64
- `installer/GridForge.iss`: Inno Setup per-user installer, shortcuts, uninstall
- `.github/workflows/windows-desktop-release.yml`: Windows runner and downloadable release artifacts

See [`docs/DESKTOP_BUILD_AND_RELEASE.md`](../docs/DESKTOP_BUILD_AND_RELEASE.md) for the build and distribution workflow, and [`docs/DESKTOP_USER_GUIDE.md`](../docs/DESKTOP_USER_GUIDE.md) for end-user installation.


