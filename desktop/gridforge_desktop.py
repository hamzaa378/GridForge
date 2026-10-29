"""GridForge Windows desktop shell.

The spreadsheet engine remains browser-side TypeScript. pywebview hosts the local,
loopback-only built site in WebView2. The installer packages this with PyInstaller.
No network connection is required for spreadsheet use.
"""
from __future__ import annotations

import base64
import binascii
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
import json
import os
from pathlib import Path
import re
import sys
import threading
import traceback
from typing import Any

APP_NAME = "GridForge"
APP_VERSION = "11.26.0"
# Bind an ephemeral loopback port for each instance. This avoids collisions with
# stale GridForge processes without exposing the local site on the network.
PORT = 0
MAX_SAVE_BYTES = 80 * 1024 * 1024
ALLOWED_EXTENSIONS = {".xlsx", ".xlsm", ".csv", ".json", ".svg"}
REQUIRED = ("index.html", "main.js", "styles.css", "engine/workbook.js")
_desktop_window: Any | None = None


def resource_dir() -> Path:
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))


def app_data_dir() -> Path:
    base = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
    return base / "GridForge"


def validate_dist(directory: Path) -> None:
    missing = [p for p in REQUIRED if not (directory / p).is_file()]
    if missing:
        raise FileNotFoundError("Missing application assets: " + ", ".join(missing))


def safe_suggested_name(name: str) -> str:
    if not isinstance(name, str):
        raise ValueError("Invalid suggested filename")
    clean = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip(" .")[:160]
    if not clean:
        clean = "GridForge.xlsx"
    if Path(clean).suffix.lower() not in ALLOWED_EXTENSIONS:
        raise ValueError("Only XLSX, XLSM, CSV, JSON and SVG files can be saved")
    return clean


class StaticHandler(SimpleHTTPRequestHandler):
    extensions_map = {
        **SimpleHTTPRequestHandler.extensions_map,
        ".js": "text/javascript; charset=utf-8",
        ".css": "text/css; charset=utf-8",
        ".html": "text/html; charset=utf-8",
        ".json": "application/json; charset=utf-8",
    }

    def end_headers(self) -> None:
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("X-Frame-Options", "DENY")
        super().end_headers()

    def list_directory(self, path: str) -> None:
        self.send_error(404, "Directory listing disabled")
        return None


class LocalServer(ThreadingHTTPServer):
    allow_reuse_address = False  # don't open an unrelated web app if already running
    daemon_threads = True


def create_local_server(dist: Path, port: int = PORT) -> LocalServer:
    validate_dist(dist)
    if not 0 <= port <= 65535:
        raise ValueError("Invalid loopback port")
    return LocalServer(("127.0.0.1", port), partial(StaticHandler, directory=str(dist)))


class DesktopAPI:
    """Only exposed native action: an explicit, user-approved Save As dialog.

    Do not expose arbitrary read/write APIs to web content. Opening files is
    already handled by the ordinary OS-backed HTML file picker.
    """

    def save_file(self, filename: str, mime: str, base64_content: str) -> str:
        import webview

        name = safe_suggested_name(filename)
        if not isinstance(base64_content, str) or len(base64_content) > (MAX_SAVE_BYTES * 4 // 3 + 16):
            raise ValueError("Workbook too large to save")
        try:
            data = base64.b64decode(base64_content, validate=True)
        except (ValueError, binascii.Error) as exc:
            raise ValueError("Invalid workbook data") from exc
        if len(data) > MAX_SAVE_BYTES:
            raise ValueError("Workbook too large to save")
        suffix = Path(name).suffix.lower()
        if suffix in (".xlsx", ".xlsm") and not data.startswith(b"PK"):
            raise ValueError("Expected an Office ZIP workbook")
        if suffix == ".svg":
            svg = data.lstrip().lower()
            if not svg.startswith(b"<svg") or any(token in svg for token in (b"<script", b"onload=", b"<foreignobject", b"javascript:")):
                raise ValueError("Only static SVG chart exports are permitted")
        masks = {
            ".xlsx": "Excel workbook (*.xlsx)",
            ".xlsm": "Macro-enabled Excel workbook (*.xlsm)",
            ".csv": "Comma-separated values (*.csv)",
            ".json": "GridForge workbook (*.json)",
            ".svg": "SVG chart (*.svg)",
        }
        window = _desktop_window
        if window is None:
            raise RuntimeError("Desktop window not ready")
        chosen = window.create_file_dialog(
            webview.FileDialog.SAVE,
            save_filename=name,
            file_types=(masks[suffix],),
        )
        if not chosen:
            return "cancelled"
        path = Path(chosen[0] if isinstance(chosen, (tuple, list)) else chosen)
        if path.suffix.lower() != suffix:
            raise ValueError("Selected file must use the expected " + suffix + " extension")
        # Use an atomic replacement to avoid incomplete workbooks on interruption.
        import tempfile
        with tempfile.NamedTemporaryFile(dir=path.parent, prefix=".gridforge-", suffix=".tmp", delete=False) as fp:
            temporary = Path(fp.name)
            fp.write(data)
            fp.flush()
            os.fsync(fp.fileno())
        try:
            os.replace(temporary, path)
        finally:
            temporary.unlink(missing_ok=True)
        return "saved"


def launch() -> int:
    import webview

    dist = resource_dir() / "dist"
    validate_dist(dist)
    settings = app_data_dir()
    settings.mkdir(parents=True, exist_ok=True)
    try:
        server = create_local_server(dist)
    except OSError as exc:
        raise RuntimeError(
            "Cannot start the GridForge desktop server on the loopback interface. "
            "Close another GridForge instance and retry."
        ) from exc
    port = server.server_address[1]
    thread = threading.Thread(target=server.serve_forever, name="gridforge-local-http", daemon=True)
    thread.start()
    api = DesktopAPI()
    try:
        window = webview.create_window(
            f"{APP_NAME} — Spreadsheet",
            url=f"http://127.0.0.1:{port}/?macro-workbench=1" if "--macro-workbench" in sys.argv else f"http://127.0.0.1:{port}",
            width=1360,
            height=850,
            min_size=(850, 550),
            js_api=api,
        )
        # Keep the native window outside the API object. Storing it on the API
        # creates a window → API → window cycle in pywebview's JS bridge, which
        # can recurse indefinitely in frozen Windows builds.
        global _desktop_window
        _desktop_window = window
        # A dedicated persistent profile + stable origin preserve workbook drafts.
        webview.start(gui="edgechromium", private_mode=False, storage_path=str(settings / "WebView2"))
    finally:
        server.shutdown()
        server.server_close()
        thread.join(timeout=3)
    return 0


def main() -> int:
    if "--check" in sys.argv:
        validate_dist(resource_dir() / "dist")
        print("PASS: bundled site and desktop launcher are present")
        return 0
    try:
        return launch()
    except Exception as exc:
        message = f"GridForge could not start.\n\n{exc}\n\nSee the desktop setup guide."
        try:
            import ctypes
            if sys.platform == "win32":
                ctypes.windll.user32.MessageBoxW(None, message, "GridForge startup error", 0x10)
            else:
                print(message, file=sys.stderr)
        except Exception:
            print(message, file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())

