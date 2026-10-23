#!/usr/bin/env python3
"""Optional LOCAL Excel/VBA bridge for GridForge.

Only Windows + desktop Microsoft Excel + pywin32 can execute Excel VBA.
Nothing executes until the user launches this service and enters its random token
into the GridForge UI and explicitly confirms one macro run.
Never expose this server on a public interface. VBA is arbitrary trusted-user code.
"""
from __future__ import annotations
import base64
import binascii
from http.server import BaseHTTPRequestHandler, HTTPServer
import io
import json
import os
import re
import secrets
import sys
import tempfile
from pathlib import Path
from zipfile import ZipFile, BadZipFile

HOST = '127.0.0.1'
PORT = int(os.environ.get('GRIDFORGE_VBA_PORT', '4187'))
TOKEN = secrets.token_urlsafe(32)
MAX_BODY = 35 * 1024 * 1024
MACRO_RE = re.compile(r'^[A-Za-z_][\w]*(?:\.[A-Za-z_][\w]*){0,3}$', re.ASCII)
LOCAL_ORIGIN = re.compile(r'^http://(?:127\.0\.0\.1|localhost):\d{1,5}$')


def validate_xlsm(raw: bytes) -> None:
    if not raw.startswith(b'PK') or len(raw) > 25 * 1024 * 1024:
        raise ValueError('Expected XLSM data up to 25 MB.')
    try:
        with ZipFile(io.BytesIO(raw)) as archive:
            names = archive.namelist()
            if len(names) > 3000 or any('../' in n or n.startswith('/') or '\\' in n for n in names):
                raise ValueError('Invalid XLSM package paths.')
            types = archive.read('[Content_Types].xml')
            if b'macroEnabled.main+xml' not in types or 'xl/workbook.xml' not in names:
                raise ValueError('Not a macro-enabled Excel workbook.')
    except (BadZipFile, KeyError) as exc:
        raise ValueError('Invalid XLSM ZIP container.') from exc


def run_excel_vba(raw: bytes, macro: str) -> bytes:
    """Run exactly one caller-selected no-argument VBA Sub through desktop Excel.

    Trust Center remains controlled by Microsoft Excel and the human operator.
    Events are disabled before loading to avoid unintended Workbook_Open handlers.
    Microsoft Excel, not GridForge, is responsible for VBA execution.
    """
    if sys.platform != 'win32':
        raise RuntimeError('VBA execution requires Windows and Microsoft Excel desktop.')
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError('Install the optional bridge dependency: py -3 -m pip install pywin32') from exc
    pythoncom.CoInitialize()
    app = None
    wb = None
    try:
        with tempfile.TemporaryDirectory(prefix='gridforge_excel_') as folder:
            path = Path(folder) / 'GridForge_Trusted_Workbook.xlsm'
            path.write_bytes(raw)
            app = win32com.client.DispatchEx('Excel.Application')
            app.Visible = True
            app.DisplayAlerts = True
            app.EnableEvents = False
            app.AskToUpdateLinks = False
            # 2 = msoAutomationSecurityByUI. Respect the user's existing Trust Center.
            # DO NOT use msoAutomationSecurityLow to bypass macro security.
            app.AutomationSecurity = 2
            wb = app.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False,
                                    IgnoreReadOnlyRecommended=True, AddToMru=False)
            # Workbook.Open should never start automatic event VBA because EnableEvents=False.
            # If Excel blocks the requested macro, let the COM exception surface to the UI.
            app.Run(f"'{wb.Name}'!{macro}")
            wb.Save()
            wb.Close(SaveChanges=False)
            wb = None
            output = path.read_bytes()
            validate_xlsm(output)
            return output
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


class Handler(BaseHTTPRequestHandler):
    server_version = 'GridForgeExcelBridge/1.0'

    def cors(self) -> bool:
        origin = self.headers.get('Origin', '')
        if origin and LOCAL_ORIGIN.fullmatch(origin):
            self.send_header('Access-Control-Allow-Origin', origin)
            self.send_header('Vary', 'Origin')
            self.send_header('Access-Control-Allow-Headers', 'Content-Type, X-GridForge-Token')
            self.send_header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS')
            return True
        return not origin

    def json_reply(self, status: int, data: dict):
        encoded = json.dumps(data).encode('utf-8')
        self.send_response(status)
        self.cors()
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Cache-Control', 'no-store')
        self.send_header('Content-Length', str(len(encoded)))
        self.end_headers()
        self.wfile.write(encoded)

    def accepted(self):
        host = self.headers.get('Host', '')
        origin = self.headers.get('Origin', '')
        return host in (f'127.0.0.1:{PORT}', f'localhost:{PORT}') and bool(LOCAL_ORIGIN.fullmatch(origin))

    def do_OPTIONS(self):
        if not self.accepted():
            self.json_reply(403, {'error': 'Only locally hosted GridForge is allowed.'})
            return
        self.send_response(204)
        self.cors()
        self.send_header('Content-Length', '0')
        self.end_headers()

    def do_GET(self):
        if self.path != '/health':
            self.json_reply(404, {'error': 'Not found.'})
        else:
            self.json_reply(200, {'status': 'ready', 'windows': sys.platform == 'win32',
                                  'excel': 'requires desktop Microsoft Excel'})

    def do_POST(self):
        if self.path != '/run':
            self.json_reply(404, {'error': 'Not found.'});return
        if not self.accepted():
            self.json_reply(403, {'error': 'Only a local browser page can request execution.'});return
        if not secrets.compare_digest(self.headers.get('X-GridForge-Token', ''), TOKEN):
            self.json_reply(403, {'error': 'Invalid session token.'});return
        try:
            size = int(self.headers.get('Content-Length', '0'))
            if size < 1 or size > MAX_BODY:
                raise ValueError('Request exceeds the configured size limit.')
            payload = json.loads(self.rfile.read(size).decode('utf-8'))
            macro = payload.get('macro')
            if not isinstance(macro, str) or not MACRO_RE.fullmatch(macro):
                raise ValueError('Use a no-argument macro such as Module1.GenerateReport.')
            encoded = payload.get('xlsm_base64')
            if not isinstance(encoded, str) or len(encoded) > MAX_BODY:
                raise ValueError('Invalid workbook encoding.')
            raw = base64.b64decode(encoded, validate=True)
            validate_xlsm(raw)
        except (ValueError, TypeError, json.JSONDecodeError, binascii.Error) as exc:
            self.json_reply(400, {'error': str(exc)});return
        try:
            result = run_excel_vba(raw, macro)
        except Exception as exc:
            self.json_reply(422, {'error': 'Excel could not run the macro: ' + str(exc)});return
        self.json_reply(200, {'status': 'executed', 'macro': macro,
                              'xlsm_base64': base64.b64encode(result).decode('ascii')})

    def log_message(self, fmt, *args):
        print('%s %s' % (self.address_string(), fmt % args), flush=True)


def main():
    if sys.platform != 'win32':
        print('This bridge can run on Windows only, with Microsoft Excel installed.', file=sys.stderr)
        sys.exit(2)
    print('\nGridForge Optional Excel VBA Bridge', flush=True)
    print('WARNING: VBA macros can access files and run other software.', flush=True)
    print(f'Local bridge: http://{HOST}:{PORT}', flush=True)
    print('Session token (paste into GridForge Automation window):', flush=True)
    print(TOKEN, flush=True)
    print('Only run macros from workbooks you trust. Ctrl+C to stop.\n', flush=True)
    HTTPServer((HOST, PORT), Handler).serve_forever()


if __name__ == '__main__':
    main()
