#!/usr/bin/env python3
"""No-dependency smoke check of the runnable GridForge release using Python stdlib."""
from __future__ import annotations

import sys
from pathlib import Path
from threading import Thread
from urllib.error import HTTPError
from urllib.request import urlopen

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from run_gridforge import create_server, validate_prebuilt  # noqa: E402


def check(condition, message):
    if not condition:
        raise AssertionError(message)
    print("PASS:", message)


def main():
    validate_prebuilt()
    check(True, "Bundled precompiled dist/ files are present")
    with create_server(0) as server:  # OS-selected available loopback port; never touches user server
        thread = Thread(target=server.serve_forever, daemon=True)
        thread.start()
        base = f"http://127.0.0.1:{server.server_port}"
        try:
            for path, expected in (
                ("/", b"GridForge"),
                ("/main.js", b""),
                ("/styles.css", b""),
                ("/engine/parser.js", b""),
                ("/engine/workbook.js", b""),
            ):
                with urlopen(base + path, timeout=5) as response:
                    content = response.read()
                    check(response.status == 200 and (expected in content if expected else len(content) > 0), f"HTTP 200 {path}")
                    if path.endswith(".js"):
                        check("javascript" in response.headers.get("Content-Type", "").lower(), f"JavaScript MIME {path}")
            try:
                urlopen(base + "/does-not-exist", timeout=5)
                raise AssertionError("Unknown path should return HTTP 404")
            except HTTPError as exc:
                check(exc.code == 404, "Missing file returns HTTP 404")
        finally:
            server.shutdown()
            thread.join(timeout=5)
    print("\nPYTHON-ONLY SMOKE CHECK: PASS. No third-party Python modules installed.")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"\nPYTHON-ONLY SMOKE CHECK FAILED: {exc}", file=sys.stderr)
        sys.exit(1)
