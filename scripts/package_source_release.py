"""Create the GitHub-ready source archive and checksums for the current release."""

from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile


ROOT = Path(__file__).resolve().parents[1]
VERSION = json.loads((ROOT / "package.json").read_text(encoding="utf-8"))["version"]
RELEASE = ROOT / "build" / "release"
PREFIX = f"GridForge-v{VERSION}-source"
ARCHIVE = RELEASE / f"{PREFIX}.zip"
GITHUB_PREFIX = f"GridForge-v{VERSION}-GitHub-Repository"
GITHUB_ARCHIVE = RELEASE / f"{GITHUB_PREFIX}.zip"
PORTABLE = ROOT / "build" / "portable"
PORTABLE_ARCHIVE = RELEASE / "GridForge-Portable-win64.zip"
INSTALLER = RELEASE / f"GridForge-Setup-{VERSION}-win64.exe"
SKIP_DIRS = {".git", "node_modules", "build", ".desktop-build-venv", "__pycache__", "work"}
SKIP_FILES = {".env", "PARITY_EXCEL_RESULTS.json", "PARITY_COMPARISON.json", "PARITY_COMPARISON.md"}


def included_files():
    for directory, names, filenames in os.walk(ROOT):
        names[:] = sorted(name for name in names if name not in SKIP_DIRS)
        folder = Path(directory)
        for filename in sorted(filenames):
            path = folder / filename
            if filename.endswith((".pyc", ".log")) or filename in SKIP_FILES:
                continue
            if folder.name == "docs" and filename in SKIP_FILES:
                continue
            yield path


def add_repository_files(archive: ZipFile, prefix: str) -> None:
    for path in included_files():
        archive.write(path, f"{prefix}/{path.relative_to(ROOT).as_posix()}")


def main() -> None:
    RELEASE.mkdir(parents=True, exist_ok=True)
    portable_app = PORTABLE / "GridForge"
    bridge = PORTABLE / "GridForgeExcelBridge.exe"
    if not portable_app.is_dir() or not bridge.is_file():
        raise SystemExit("Build the portable desktop app first with desktop/build_windows.ps1 -SkipTests -PortableOnly.")
    with ZipFile(PORTABLE_ARCHIVE, "w", compression=ZIP_DEFLATED, compresslevel=6) as archive:
        for path in sorted(portable_app.rglob("*")):
            if path.is_file() and "__pycache__" not in path.parts and path.suffix != ".pyc":
                archive.write(path, path.relative_to(PORTABLE).as_posix())
        archive.write(bridge, bridge.relative_to(PORTABLE).as_posix())

    if not INSTALLER.is_file():
        raise SystemExit(f"Build the Windows installer first; missing {INSTALLER.name}.")
    manifest_path = ROOT / "docs" / "DESKTOP_RELEASE_MANIFEST.json"
    if manifest_path.is_file():
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        for key, path in (("installer", INSTALLER), ("portable", PORTABLE_ARCHIVE)):
            manifest.setdefault("artifacts", {})[key] = {
                "file": path.name,
                "size_bytes": path.stat().st_size,
                "sha256": hashlib.sha256(path.read_bytes()).hexdigest().upper(),
            }
        manifest_path.write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")

    with ZipFile(ARCHIVE, "w", compression=ZIP_DEFLATED, compresslevel=9) as archive:
        add_repository_files(archive, PREFIX)
    with ZipFile(GITHUB_ARCHIVE, "w", compression=ZIP_DEFLATED, compresslevel=6) as archive:
        add_repository_files(archive, GITHUB_PREFIX)
        installer_member = f"{GITHUB_PREFIX}/release/{INSTALLER.name}"
        archive.write(INSTALLER, installer_member)
        installer_digest = hashlib.sha256(INSTALLER.read_bytes()).hexdigest().upper()
        archive.writestr(
            f"{GITHUB_PREFIX}/release/SHA256SUMS.txt",
            f"{installer_digest}  {INSTALLER.name}\n",
        )

    artifacts = [
        INSTALLER,
        PORTABLE_ARCHIVE,
        ARCHIVE,
        GITHUB_ARCHIVE,
    ]
    missing = [path.name for path in artifacts if not path.is_file()]
    if missing:
        raise SystemExit(f"Required release artifacts are missing: {', '.join(missing)}")
    lines = []
    for path in artifacts:
        digest = hashlib.sha256(path.read_bytes()).hexdigest().upper()
        lines.append(f"{digest}  {path.name}")
    (RELEASE / "SHA256SUMS.txt").write_text("\n".join(lines) + "\n", encoding="ascii")
    print(f"Created {ARCHIVE}")
    print(f"Created {GITHUB_ARCHIVE}")
    print(f"Updated {RELEASE / 'SHA256SUMS.txt'}")


if __name__ == "__main__":
    main()
