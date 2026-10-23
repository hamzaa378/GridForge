#!/usr/bin/env python3
"""Run the INCLUDED precompiled GridForge site using only Python's standard library.

No Node.js, npm, pip, backend, database or network downloads are needed to RUN.
To EDIT/BUILD TypeScript or run JS unit tests, install Node.js separately.
"""
from __future__ import annotations

import argparse
from functools import partial
import http.server
import os
from pathlib import Path
import sys
import webbrowser

ROOT = Path(__file__).resolve().parent
DIST = ROOT / "dist"
REQUIRED = ("index.html", "main.js", "styles.css", "engine/parser.js", "engine/workbook.js")


class GridForgeHandler(http.server.SimpleHTTPRequestHandler):
    extensions_map = {
        **http.server.SimpleHTTPRequestHandler.extensions_map,
        ".js": "text/javascript; charset=utf-8",
        ".css": "text/css; charset=utf-8",
        ".html": "text/html; charset=utf-8",
        ".json": "application/json; charset=utf-8",
    }

    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        super().end_headers()

    def list_directory(self, path):
        self.send_error(404, "Directory browsing disabled")
        return None


def validate_prebuilt():
    missing = [name for name in REQUIRED if not (DIST / name).is_file()]
    if missing:
        raise FileNotFoundError(
            "Prebuilt web application files are missing: " + ", ".join(missing)
            + ". Download/extract the COMPLETE release ZIP, including dist/. "
            + "Developers can regenerate dist/ with npm install && npm run build."
        )


def create_server(port: int):
    handler = partial(GridForgeHandler, directory=str(DIST))
    return http.server.ThreadingHTTPServer(("127.0.0.1", port), handler)


def main(argv=None):
    parser = argparse.ArgumentParser(description="GridForge: run without Node.js or pip packages")
    parser.add_argument("--port", type=int, default=None, help="Preferred local port (default: 4174)")
    parser.add_argument("--no-browser", action="store_true", help="Don't automatically open a browser")
    parser.add_argument("--check", action="store_true", help="Verify bundled web files without starting a server")
    args = parser.parse_args(argv)
    try:
        validate_prebuilt()
    except FileNotFoundError as exc:
        print(f"SETUP ERROR: {exc}", file=sys.stderr)
        return 2
    if args.check:
        print("PASS: Bundled GridForge web application is present; Python-only launch is available.")
        return 0
    try:
        preferred_port = args.port if args.port is not None else int(os.getenv("PORT", "4174"))
    except ValueError:
        parser.error("PORT must be an integer")
    if not 1 <= preferred_port <= 65535:
        parser.error("PORT must be between 1 and 65535")
    server = None
    # Match the Node launcher: try a few nearby ports unless explicitly selected.
    port_range = [preferred_port] if args.port is not None or "PORT" in os.environ else list(range(4174, 4185))
    last_error = None
    for port in port_range:
        try:
            server = create_server(port)
            break
        except OSError as exc:
            last_error = exc
            if exc.errno not in (48, 98, 10048):  # address already in use on macOS/Linux/Windows
                print(f"Unable to bind loopback server: {exc}", file=sys.stderr)
                return 1
    if server is None:
        print(f"No free local port in {port_range}: {last_error}", file=sys.stderr)
        print("Try python run_gridforge.py --port 4190", file=sys.stderr)
        return 1
    url = f"http://127.0.0.1:{server.server_port}"
    print("\nGridForge spreadsheet — Python-only local launcher", flush=True)
    print(f"Open in your browser: {url}", flush=True)
    print("No pip or npm installation is required to run this packaged version.", flush=True)
    print("Keep this terminal open. Press Ctrl+C to stop the server.\n", flush=True)
    if not args.no_browser:
        webbrowser.open(url)
    try:
        server.serve_forever(poll_interval=0.2)
    except KeyboardInterrupt:
        print("\nGridForge server stopped.")
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
