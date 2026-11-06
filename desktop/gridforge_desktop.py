"""GridForge Windows desktop shell.

The spreadsheet engine remains browser-side TypeScript. pywebview hosts the local,
loopback-only built site in WebView2. The installer packages this with PyInstaller.
Local spreadsheet use works offline. Optional cloud features require network access.
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
import secrets
import sqlite3
import sys
import threading
import time
import traceback
from typing import Any
from urllib.parse import parse_qs, urlencode, urlparse
from urllib.error import HTTPError, URLError
from urllib.request import HTTPRedirectHandler, Request, build_opener, urlopen

APP_NAME = "GridForge"
APP_VERSION = "12.0.210"
# Bind an ephemeral loopback port for each instance. This avoids collisions with
# stale GridForge processes without exposing the local site on the network.
PORT = 0
MAX_SAVE_BYTES = 80 * 1024 * 1024
ALLOWED_EXTENSIONS = {".xlsx", ".xlsm", ".ods", ".csv", ".json", ".svg"}
REQUIRED = ("index.html", "main.js", "styles.css", "engine/workbook.js")
_desktop_window: Any | None = None
_cloud_auth_callbacks: dict[str, tuple[Any, threading.Event, dict[str, str]]] = {}
_cloud_auth_lock = threading.Lock()
_CLOUD_AUTH_HOSTS = {"onedrive": "login.microsoftonline.com", "googledrive": "accounts.google.com", "dropbox": "www.dropbox.com", "box": "account.box.com"}


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
        raise ValueError("Only XLSX, XLSM, ODS, CSV, JSON and SVG files can be saved")
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
    """Expose narrowly scoped native file, query, and cloud actions to the UI.

    Keep methods bounded and explicit; do not expose arbitrary filesystem or
    process APIs to web content. Secrets should never enter workbook state.
    """

    def __init__(self) -> None:
        self._sqlite_lock = threading.Lock()
        self._sqlite_handles: dict[str, Path] = {}

    def choose_sqlite_database(self) -> dict[str, str] | None:
        """Ask the user to select a local SQLite file and return an opaque session handle."""
        import webview

        window = _desktop_window
        if window is None:
            raise RuntimeError("Desktop window not ready")
        chosen = window.create_file_dialog(
            webview.FileDialog.OPEN,
            allow_multiple=False,
            file_types=("SQLite databases (*.db;*.sqlite;*.sqlite3)", "All files (*.*)"),
        )
        if not chosen:
            return None
        path = Path(chosen[0] if isinstance(chosen, (tuple, list)) else chosen).resolve(strict=True)
        if not path.is_file() or path.suffix.lower() not in {".db", ".sqlite", ".sqlite3"}:
            raise ValueError("Select a SQLite database file (.db, .sqlite or .sqlite3)")
        handle = secrets.token_urlsafe(24)
        with self._sqlite_lock:
            self._sqlite_handles[handle] = path
            if len(self._sqlite_handles) > 24:
                self._sqlite_handles.pop(next(iter(self._sqlite_handles)))
        return {"handle": handle, "name": path.name}

    def run_sqlite_select(self, handle: str, sql: str) -> dict[str, Any]:
        """Execute one bounded read-only SELECT against a database chosen in this session."""
        if not isinstance(handle, str) or len(handle) > 128 or not isinstance(sql, str) or not 1 <= len(sql.strip()) <= 20_000 or "\x00" in sql:
            raise ValueError("Enter one SQL query of at most 20,000 characters")
        if not re.match(r"^\s*(?:SELECT|WITH)\b", sql, re.IGNORECASE):
            raise ValueError("Only read-only SELECT queries are allowed")
        with self._sqlite_lock:
            path = self._sqlite_handles.get(handle)
        if path is None:
            raise ValueError("The selected database session expired. Choose the database file again.")
        if not path.is_file():
            raise ValueError("The selected database file is no longer available")
        uri = path.as_uri() + "?mode=ro"
        connection: sqlite3.Connection | None = None
        deadline = time.monotonic() + 8.0
        instruction_count = 0
        try:
            connection = sqlite3.connect(uri, uri=True, timeout=2.0)
            connection.enable_load_extension(False)
            connection.execute("PRAGMA query_only=ON")
            connection.execute("PRAGMA trusted_schema=OFF")
            limits = (
                ("SQLITE_LIMIT_LENGTH", 32767), ("SQLITE_LIMIT_SQL_LENGTH", 20_000),
                ("SQLITE_LIMIT_COLUMN", 104), ("SQLITE_LIMIT_EXPR_DEPTH", 64),
                ("SQLITE_LIMIT_COMPOUND_SELECT", 32), ("SQLITE_LIMIT_VARIABLE_NUMBER", 1000),
                ("SQLITE_LIMIT_ATTACHED", 0),
            )
            for name, limit in limits:
                category = getattr(sqlite3, name, None)
                if category is not None:
                    connection.setlimit(category, limit)
            denied_names = (
                "SQLITE_INSERT", "SQLITE_UPDATE", "SQLITE_DELETE", "SQLITE_CREATE_INDEX", "SQLITE_CREATE_TABLE",
                "SQLITE_CREATE_TEMP_INDEX", "SQLITE_CREATE_TEMP_TABLE", "SQLITE_CREATE_TEMP_TRIGGER",
                "SQLITE_CREATE_TEMP_VIEW", "SQLITE_CREATE_TRIGGER", "SQLITE_CREATE_VIEW", "SQLITE_DROP_INDEX",
                "SQLITE_DROP_TABLE", "SQLITE_DROP_TEMP_INDEX", "SQLITE_DROP_TEMP_TABLE", "SQLITE_DROP_TEMP_TRIGGER",
                "SQLITE_DROP_TEMP_VIEW", "SQLITE_DROP_TRIGGER", "SQLITE_DROP_VIEW", "SQLITE_ALTER_TABLE",
                "SQLITE_ATTACH", "SQLITE_DETACH", "SQLITE_PRAGMA", "SQLITE_TRANSACTION", "SQLITE_SAVEPOINT",
                "SQLITE_CREATE_VTABLE", "SQLITE_DROP_VTABLE", "SQLITE_REINDEX", "SQLITE_ANALYZE",
            )
            denied = {getattr(sqlite3, name) for name in denied_names if hasattr(sqlite3, name)}

            def authorize(action: int, first: str | None, second: str | None, database: str | None, trigger: str | None) -> int:
                if action in denied:
                    return sqlite3.SQLITE_DENY
                if action == getattr(sqlite3, "SQLITE_FUNCTION", -1):
                    function_name = (second or first or "").lower()
                    if function_name in {"load_extension", "readfile", "writefile", "fts3_tokenizer"}:
                        return sqlite3.SQLITE_DENY
                return sqlite3.SQLITE_OK

            connection.set_authorizer(authorize)

            def progress() -> int:
                nonlocal instruction_count
                instruction_count += 1000
                return int(time.monotonic() > deadline or instruction_count > 5_000_000)

            connection.set_progress_handler(progress, 1000)
            cursor = connection.execute(sql)
            description = cursor.description
            if not description:
                raise ValueError("The SQL query must return a result table")
            if len(description) > 104:
                raise ValueError("SQLite results are limited to 104 columns")
            headers: list[str] = []
            counts: dict[str, int] = {}
            for index, column in enumerate(description, 1):
                base = str(column[0] or f"Column {index}")[:255]
                count = counts.get(base.casefold(), 0) + 1
                counts[base.casefold()] = count
                headers.append(base if count == 1 else f"{base} ({count})")
            rows: list[list[str | int | float | bool | None]] = []
            for row in cursor.fetchmany(1001):
                if len(rows) >= 1000:
                    raise ValueError("SQLite results are limited to 1,000 data rows")
                clean: list[str | int | float | bool | None] = []
                for cell in row:
                    if cell is None or isinstance(cell, (str, bool)):
                        if isinstance(cell, str) and len(cell) > 32_767:
                            raise ValueError("SQLite text cells are limited to 32,767 characters")
                        clean.append(cell)
                    elif isinstance(cell, int):
                        clean.append(cell if abs(cell) <= 9_007_199_254_740_991 else str(cell))
                    elif isinstance(cell, float) and (cell == cell) and abs(cell) != float("inf"):
                        clean.append(cell)
                    else:
                        raise ValueError("SQLite BLOB or non-finite values cannot be loaded into worksheet cells")
                rows.append(clean)
                if (len(rows) + 1) * len(headers) > 12_000:
                    raise ValueError("SQLite results are limited to 12,000 total cells")
            return {"headers": headers, "rows": rows, "database": path.name}
        except sqlite3.Error as exc:
            message = "The SQL query exceeded its time limit" if time.monotonic() > deadline else str(exc)
            raise ValueError(message[:500]) from exc
        finally:
            if connection is not None:
                connection.close()

    def fetch_authenticated_data(self, url: str, kind: str, auth_kind: str, username: str = "", secret: str = "") -> dict[str, str]:
        """Fetch a bounded HTTPS CSV/JSON document using credentials for this request only."""
        if not isinstance(url, str) or len(url) > 2048 or "\r" in url or "\n" in url:
            raise ValueError("Enter a valid HTTPS data URL")
        parsed = urlparse(url)
        if parsed.scheme != "https" or not parsed.hostname or parsed.username or parsed.password or parsed.fragment:
            raise ValueError("Authenticated data imports require an HTTPS URL without embedded credentials or fragments")
        if kind not in {"csv", "json"}:
            raise ValueError("Choose CSV or JSON as the response format")
        if auth_kind not in {"none", "basic", "bearer"}:
            raise ValueError("Choose no authentication, Basic, or Bearer authentication")
        if not isinstance(username, str) or not isinstance(secret, str) or len(username) > 512 or len(secret) > 8192:
            raise ValueError("Credentials exceed the supported input length")
        if any(char in username + secret for char in "\r\n\x00"):
            raise ValueError("Credentials cannot contain line breaks or null characters")
        headers = {"Accept": "text/csv, text/plain;q=0.9" if kind == "csv" else "application/json"}
        if auth_kind == "basic":
            if not username or not secret or ":" in username:
                raise ValueError("Basic authentication requires a username and password")
            token = base64.b64encode(f"{username}:{secret}".encode("utf-8")).decode("ascii")
            headers["Authorization"] = f"Basic {token}"
        elif auth_kind == "bearer":
            if not re.fullmatch(r"[A-Za-z0-9._~+/-]+=*", secret):
                raise ValueError("Bearer authentication requires a token")
            headers["Authorization"] = f"Bearer {secret}"
        request = Request(url, headers=headers, method="GET")

        class RejectRedirects(HTTPRedirectHandler):
            def redirect_request(self, req: Request, fp: Any, code: int, msg: str, response_headers: Any, new_url: str) -> None:
                return None

        try:
            with build_opener(RejectRedirects).open(request, timeout=15) as response:
                length = response.headers.get("Content-Length")
                if length and int(length) > 3 * 1024 * 1024:
                    raise ValueError("External data is limited to 3 MB")
                content = response.read(3 * 1024 * 1024 + 1)
                if len(content) > 3 * 1024 * 1024:
                    raise ValueError("External data is limited to 3 MB")
                try:
                    text = content.decode("utf-8-sig", errors="strict")
                except UnicodeDecodeError as exc:
                    raise ValueError("External data must be valid UTF-8 text") from exc
                return {"content": text, "status": str(response.status)}
        except HTTPError as exc:
            if 300 <= exc.code < 400:
                raise ValueError("Redirects are not followed for authenticated requests") from exc
            if exc.code in {401, 403}:
                raise ValueError("The server rejected these credentials (HTTP 401/403)") from exc
            raise ValueError(f"The data source returned HTTP {exc.code}") from exc
        except URLError as exc:
            raise ValueError("The HTTPS data source could not be reached") from exc

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
        if suffix in (".xlsx", ".xlsm", ".ods") and not data.startswith(b"PK"):
            raise ValueError("Expected an Office ZIP workbook")
        if suffix == ".svg":
            svg = data.lstrip().lower()
            if not svg.startswith(b"<svg") or any(token in svg for token in (b"<script", b"onload=", b"<foreignobject", b"javascript:")):
                raise ValueError("Only static SVG chart exports are permitted")
        masks = {
            ".xlsx": "Excel workbook (*.xlsx)",
            ".xlsm": "Macro-enabled Excel workbook (*.xlsm)",
            ".ods": "OpenDocument spreadsheet (*.ods)",
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

    def start_cloud_auth(self, provider: str, dropbox_redirect_uri: str = "") -> dict[str, str]:
        """Start a short-lived, loopback-only OAuth callback listener."""
        from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

        if provider not in _CLOUD_AUTH_HOSTS:
            raise ValueError("Unsupported cloud provider")
        auth_id = secrets.token_urlsafe(24)
        completed = threading.Event()
        result: dict[str, str] = {}

        class CallbackHandler(BaseHTTPRequestHandler):
            def do_GET(self) -> None:
                parsed = urlparse(self.path)
                if parsed.path != "/oauth/callback":
                    self.send_error(404)
                    return
                values = parse_qs(parsed.query, keep_blank_values=True)
                for key in ("code", "state", "error", "error_description"):
                    value = (values.get(key) or [""])[0]
                    if len(value) > 4096:
                        self.send_error(400)
                        return
                    if value:
                        result[key] = value
                completed.set()
                body = b"<!doctype html><meta charset=utf-8><title>GridForge sign-in</title><p>Sign-in response received. You can return to GridForge.</p>"
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(body)))
                self.send_header("Cache-Control", "no-store")
                self.send_header("X-Content-Type-Options", "nosniff")
                self.end_headers()
                self.wfile.write(body)

            def log_message(self, _format: str, *_args: Any) -> None:
                return

        port = 0
        if provider in {"dropbox", "box"}:
            parsed = urlparse(dropbox_redirect_uri)
            if parsed.scheme != "http" or parsed.hostname != "localhost" or parsed.path != "/oauth/callback" or not parsed.port or parsed.username or parsed.password:
                raise ValueError(f"Register the exact localhost redirect URI in the {'Box' if provider == 'box' else 'Dropbox'} app console.")
            port = parsed.port
        server = ThreadingHTTPServer(("127.0.0.1", port), CallbackHandler)
        server.daemon_threads = True
        thread = threading.Thread(target=server.serve_forever, name="gridforge-cloud-oauth", daemon=True)
        thread.start()
        redirect_host = "localhost" if provider in {"onedrive", "dropbox", "box"} else "127.0.0.1"
        redirect_uri = f"http://{redirect_host}:{server.server_address[1]}/oauth/callback"
        with _cloud_auth_lock:
            stale = list(_cloud_auth_callbacks)
            for key in stale:
                old_server, _old_event, _old_result = _cloud_auth_callbacks.pop(key)
                old_server.shutdown()
                old_server.server_close()
            _cloud_auth_callbacks[auth_id] = (server, completed, result)
        return {"id": auth_id, "redirectUri": redirect_uri}

    def open_cloud_auth(self, provider: str, url: str) -> bool:
        import webbrowser

        parsed = urlparse(url)
        if provider not in _CLOUD_AUTH_HOSTS or parsed.scheme != "https" or parsed.hostname != _CLOUD_AUTH_HOSTS[provider] or parsed.username or parsed.password:
            raise ValueError("Cloud sign-in URL was rejected")
        return bool(webbrowser.open(url, new=2, autoraise=True))

    def poll_cloud_auth(self, auth_id: str, timeout_seconds: float = 1.0) -> dict[str, str] | None:
        if not isinstance(auth_id, str) or len(auth_id) > 128:
            raise ValueError("Invalid sign-in session")
        with _cloud_auth_lock:
            entry = _cloud_auth_callbacks.get(auth_id)
        if not entry:
            return None
        server, completed, result = entry
        if not completed.wait(max(0.0, min(2.0, float(timeout_seconds)))):
            return None
        with _cloud_auth_lock:
            _cloud_auth_callbacks.pop(auth_id, None)
        server.shutdown()
        server.server_close()
        return dict(result)

    def cancel_cloud_auth(self, auth_id: str) -> bool:
        if not isinstance(auth_id, str) or len(auth_id) > 128:
            return False
        with _cloud_auth_lock:
            entry = _cloud_auth_callbacks.pop(auth_id, None)
        if not entry:
            return False
        server, _completed, _result = entry
        server.shutdown()
        server.server_close()
        return True

    def exchange_cloud_code(self, provider: str, client_id: str, code: str, verifier: str, redirect_uri: str, client_secret: str = "") -> dict[str, Any]:
        endpoints = {
            "onedrive": "https://login.microsoftonline.com/common/oauth2/v2.0/token",
            "googledrive": "https://oauth2.googleapis.com/token",
            "dropbox": "https://api.dropbox.com/oauth2/token",
            "box": "https://api.box.com/oauth2/token",
        }
        if provider not in endpoints or not isinstance(client_id, str) or not 1 <= len(client_id) <= 512 or not isinstance(code, str) or not 1 <= len(code) <= 4096 or not isinstance(verifier, str) or (provider != "box" and (not 43 <= len(verifier) <= 128 or not re.fullmatch(r"[A-Za-z0-9._~-]+", verifier))) or (provider == "box" and verifier):
            raise ValueError("Invalid cloud sign-in data")
        parsed = urlparse(redirect_uri)
        if parsed.scheme != "http" or parsed.hostname not in {"localhost", "127.0.0.1"} or parsed.path != "/oauth/callback" or not parsed.port:
            raise ValueError("Cloud sign-in callback was rejected")
        fields = {"client_id": client_id, "grant_type": "authorization_code", "code": code, "redirect_uri": redirect_uri}
        if provider != "box":
            fields["code_verifier"] = verifier
        if provider in {"googledrive", "box"}:
            if provider == "box" and (not isinstance(client_secret, str) or not client_secret or len(client_secret) > 512):
                raise ValueError("Enter the Box app client secret.")
            if client_secret:
                fields["client_secret"] = client_secret
        request = Request(endpoints[provider], data=urlencode(fields).encode(), headers={"Content-Type": "application/x-www-form-urlencoded", "Accept": "application/json"}, method="POST")
        try:
            with urlopen(request, timeout=30) as response:
                payload = response.read(131073)
        except HTTPError as exc:
            detail = exc.read(65536).decode("utf-8", "replace")
            try:
                body = json.loads(detail)
                detail = str(body.get("error_description") or body.get("error") or "Sign-in was rejected")
            except (ValueError, AttributeError):
                detail = "Sign-in was rejected"
            raise ValueError(detail[:500]) from exc
        except (URLError, TimeoutError) as exc:
            raise RuntimeError("Could not reach the cloud sign-in service") from exc
        if len(payload) > 131072:
            raise ValueError("Cloud sign-in response was too large")
        try:
            token = json.loads(payload)
        except (ValueError, UnicodeDecodeError) as exc:
            raise ValueError("Cloud sign-in returned an invalid response") from exc
        if not isinstance(token, dict) or not isinstance(token.get("access_token"), str):
            raise ValueError(str(token.get("error_description") or token.get("error") or "Cloud sign-in did not return an access token")[:500])
        return {key: token[key] for key in ("access_token", "refresh_token", "expires_in", "scope", "token_type") if key in token}

    def refresh_cloud_token(self, provider: str, client_id: str, refresh_token: str, client_secret: str = "") -> dict[str, Any]:
        endpoints = {"onedrive": "https://login.microsoftonline.com/common/oauth2/v2.0/token", "googledrive": "https://oauth2.googleapis.com/token", "dropbox": "https://api.dropbox.com/oauth2/token", "box": "https://api.box.com/oauth2/token"}
        if provider not in endpoints or not isinstance(client_id, str) or not client_id or len(client_id) > 512 or not isinstance(refresh_token, str) or not refresh_token or len(refresh_token) > 8192:
            raise ValueError("Invalid cloud refresh data")
        fields = {"client_id": client_id, "grant_type": "refresh_token", "refresh_token": refresh_token}
        if provider in {"googledrive", "box"}:
            if provider == "box" and not client_secret:
                raise ValueError("The protected Box app client secret is missing. Re-enter it and connect again.")
            if client_secret:
                fields["client_secret"] = client_secret
        request = Request(endpoints[provider], data=urlencode(fields).encode(), headers={"Content-Type": "application/x-www-form-urlencoded", "Accept": "application/json"}, method="POST")
        try:
            with urlopen(request, timeout=30) as response:
                payload = response.read(131073)
        except (HTTPError, URLError, TimeoutError) as exc:
            raise RuntimeError("The saved cloud sign-in expired or could not be refreshed. Reconnect this account.") from exc
        if len(payload) > 131072:
            raise ValueError("Cloud refresh response was too large")
        token = json.loads(payload)
        if not isinstance(token, dict) or not isinstance(token.get("access_token"), str):
            raise ValueError("Cloud refresh did not return an access token")
        return {key: token[key] for key in ("access_token", "refresh_token", "expires_in", "scope", "token_type") if key in token}

    def cloud_credential(self, provider: str, value: str | None = None) -> str | None:
        if provider not in _CLOUD_AUTH_HOSTS:
            raise ValueError("Unsupported cloud provider")
        import win32crypt

        directory = app_data_dir() / "CloudCredentials"
        path = directory / (provider + ".bin")
        if value is None:
            if not path.is_file():
                return None
            protected = path.read_bytes()
            clear = win32crypt.CryptUnprotectData(protected, None, None, None, 1)[1]
            return clear.decode("utf-8")
        if not isinstance(value, str) or len(value) > 16384:
            raise ValueError("Cloud credential data exceeded its limit")
        directory.mkdir(parents=True, exist_ok=True)
        protected = win32crypt.CryptProtectData(value.encode("utf-8"), "GridForge cloud credential", None, None, None, 1)
        import tempfile
        with tempfile.NamedTemporaryFile(dir=directory, prefix=".credential-", suffix=".tmp", delete=False) as fp:
            temporary = Path(fp.name)
            fp.write(protected)
            fp.flush()
            os.fsync(fp.fileno())
        try:
            os.replace(temporary, path)
        finally:
            temporary.unlink(missing_ok=True)
        return "saved"

    def box_client_secret(self, value: str | None = None) -> str | None:
        """Read or replace the Box OAuth client secret protected with Windows DPAPI."""
        if value is None:
            path = app_data_dir() / "CloudCredentials" / "box-client-secret.bin"
            if not path.is_file():
                return None
            import win32crypt
            protected = path.read_bytes()
            return win32crypt.CryptUnprotectData(protected, None, None, None, 1)[1].decode("utf-8")
        if not isinstance(value, str) or len(value) > 512:
            raise ValueError("Invalid Box client secret")
        path = app_data_dir() / "CloudCredentials" / "box-client-secret.bin"
        if not value:
            path.unlink(missing_ok=True)
            return None
        import win32crypt
        path.parent.mkdir(parents=True, exist_ok=True)
        protected = win32crypt.CryptProtectData(value.encode("utf-8"), "GridForge Box app secret", None, None, None, 1)
        import tempfile
        with tempfile.NamedTemporaryFile(dir=path.parent, prefix=".box-secret-", suffix=".tmp", delete=False) as fp:
            temporary = Path(fp.name)
            fp.write(protected)
            fp.flush()
            os.fsync(fp.fileno())
        try:
            os.replace(temporary, path)
        finally:
            temporary.unlink(missing_ok=True)
        return "saved"

    def box_api_request(self, method: str, path: str, access_token: str, body_base64: str = "", content_type: str = "application/json", if_match: str = "") -> dict[str, Any]:
        """Call the small Box API surface used by cloud workbooks from native code.

        Keeping these requests in the desktop process avoids Box's browser CORS
        allow-list requirement, which cannot reliably match GridForge's random
        local WebView port. Only fixed Box API hosts and workbook endpoints are
        accepted; download redirects are restricted to Box's download host.
        """
        from urllib.request import HTTPRedirectHandler, build_opener

        if method not in {"GET", "POST"} or not isinstance(access_token, str) or not 1 <= len(access_token) <= 16384:
            raise ValueError("Invalid Box request")
        if not isinstance(path, str) or len(path) > 8192 or "\\" in path or "\r" in path or "\n" in path:
            raise ValueError("Invalid Box API path")
        parsed_path = urlparse(path)
        if parsed_path.scheme or parsed_path.netloc or parsed_path.fragment:
            raise ValueError("Invalid Box API path")
        upload = parsed_path.path.startswith("/api/2.0/")
        if upload and (method != "POST" or not (parsed_path.path == "/api/2.0/files/content" or re.fullmatch(r"/api/2\.0/files/[A-Za-z0-9_-]{1,128}/content", parsed_path.path)) or parsed_path.query):
            raise ValueError("Unsupported Box upload endpoint")
        allowed_list_query = {"usemarker", "limit", "fields", "marker"}
        list_query = parse_qs(parsed_path.query, keep_blank_values=True)
        if not upload and not (re.fullmatch(r"/2\.0/files/[A-Za-z0-9_-]{1,128}/content", parsed_path.path) and method == "GET" and not parsed_path.query or parsed_path.path == "/2.0/folders/0/items" and method == "GET" and set(list_query) <= allowed_list_query and list_query.get("usemarker") == ["true"] and list_query.get("limit") == ["100"] and list_query.get("fields") == ["id,type,name,size,modified_at,etag,file_version"] and len(list_query.get("marker", [""])[0]) <= 4096):
            raise ValueError("Unsupported Box workbook endpoint")
        max_upload_bytes = 50 * 1024 * 1024 + 65536
        if not isinstance(body_base64, str) or len(body_base64) > ((max_upload_bytes + 2) // 3) * 4 + 16:
            raise ValueError("Box workbook upload exceeds 50 MB")
        try:
            payload = base64.b64decode(body_base64, validate=True) if body_base64 else b""
        except (ValueError, binascii.Error) as exc:
            raise ValueError("Invalid Box request payload") from exc
        if len(payload) > max_upload_bytes:
            raise ValueError("Box workbook upload exceeds 50 MB")
        if content_type not in {"application/json", "multipart/form-data; boundary=gridforge-box"}:
            raise ValueError("Unsupported Box content type")
        if if_match and (len(if_match) > 256 or not re.fullmatch(r"[A-Za-z0-9._-]+", if_match)):
            raise ValueError("Invalid Box file version")
        origin = "https://upload.box.com" if upload else "https://api.box.com"
        request = Request(origin + path, data=payload if method == "POST" else None, headers={"Authorization": "Bearer " + access_token, "Accept": "application/json, application/octet-stream", **({"Content-Type": content_type} if method == "POST" else {}), **({"If-Match": if_match} if if_match else {})}, method=method)

        class BoxRedirectHandler(HTTPRedirectHandler):
            def redirect_request(self, req: Request, fp: Any, code: int, msg: str, headers: Any, newurl: str) -> Request:
                parsed = urlparse(newurl)
                allowed_hosts = {"dl.boxcloud.com"} if not upload else {"upload.box.com", "api.box.com"}
                if parsed.scheme != "https" or parsed.hostname not in allowed_hosts or parsed.username or parsed.password:
                    raise ValueError("Box returned an untrusted redirect address")
                safe_headers = dict(req.header_items())
                if parsed.hostname != urlparse(req.full_url).hostname:
                    safe_headers = {key: value for key, value in safe_headers.items() if key.lower() not in {"authorization", "content-type", "if-match"}}
                return Request(newurl, headers=safe_headers, method="GET")

        opener = build_opener(BoxRedirectHandler())
        try:
            with opener.open(request, timeout=45) as response:
                data = response.read(80 * 1024 * 1024 + 1)
                if len(data) > 80 * 1024 * 1024:
                    raise ValueError("Box response exceeds GridForge's 80 MB workbook limit")
                return {"status": response.status, "contentType": response.headers.get("Content-Type", ""), "bodyBase64": base64.b64encode(data).decode("ascii")}
        except HTTPError as exc:
            detail = exc.read(65536).decode("utf-8", "replace")
            try:
                error = json.loads(detail)
                detail = str(error.get("message") or error.get("error_description") or error.get("code") or "Box request failed")
            except (ValueError, AttributeError):
                detail = "Box request failed"
            raise ValueError(detail[:500]) from exc
        except (URLError, TimeoutError) as exc:
            raise RuntimeError("Could not reach Box") from exc

    def remove_cloud_credential(self, provider: str) -> bool:
        if provider not in _CLOUD_AUTH_HOSTS:
            raise ValueError("Unsupported cloud provider")
        path = app_data_dir() / "CloudCredentials" / (provider + ".bin")
        existed = path.exists()
        path.unlink(missing_ok=True)
        if provider == "box":
            secret_path = app_data_dir() / "CloudCredentials" / "box-client-secret.bin"
            existed = secret_path.exists() or existed
            secret_path.unlink(missing_ok=True)
        return existed


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






