"""Regression tests runnable on Linux and Windows without a graphical desktop."""
import base64
from functools import partial
from http.client import HTTPConnection
from pathlib import Path
import re
import sqlite3
import sys
import tempfile
import threading
import types
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'desktop'))
import gridforge_desktop as desktop_module
from gridforge_desktop import (
    DesktopAPI, MAX_SAVE_BYTES, PORT, create_local_server, resource_dir,
    safe_suggested_name, validate_dist,
)

ROOT = Path(__file__).resolve().parents[1]


class DesktopTests(unittest.TestCase):
    def test_sqlite_connector_is_read_only_bounded_and_keeps_formula_text(self):
        with tempfile.TemporaryDirectory() as folder:
            database = Path(folder) / 'sales.sqlite'
            connection = sqlite3.connect(database)
            try:
                connection.execute('CREATE TABLE Sales (id INTEGER, name TEXT, formula TEXT, large INTEGER)')
                connection.execute('INSERT INTO Sales VALUES (?, ?, ?, ?)', (1, 'Ada', '=1+1', 9_223_372_036_854_775_807))
                connection.commit()
            finally:
                connection.close()
            api = DesktopAPI()
            api._sqlite_handles['session-handle'] = database
            result = api.run_sqlite_select('session-handle', 'SELECT id AS Key, name AS key, formula, large FROM Sales')
            self.assertEqual(result['headers'], ['Key', 'key (2)', 'formula', 'large'])
            self.assertEqual(result['rows'], [[1, 'Ada', '=1+1', '9223372036854775807']])
            self.assertEqual(result['database'], 'sales.sqlite')
            for sql in ('DELETE FROM Sales', 'PRAGMA table_info(Sales)', 'SELECT load_extension(\'missing\')', 'SELECT 1; DELETE FROM Sales'):
                with self.subTest(sql=sql):
                    with self.assertRaises(ValueError):
                        api.run_sqlite_select('session-handle', sql)
            self.assertEqual(api.run_sqlite_select('session-handle', 'SELECT COUNT(*) FROM Sales')['rows'], [[1]])
            with self.assertRaisesRegex(ValueError, 'session expired'):
                api.run_sqlite_select('unknown-handle', 'SELECT 1')

    def test_sqlite_connector_caps_returned_rows(self):
        with tempfile.TemporaryDirectory() as folder:
            database = Path(folder) / 'rows.db'
            api = DesktopAPI()
            api._sqlite_handles['rows'] = database
            connection = sqlite3.connect(database)
            try:
                connection.execute('CREATE TABLE Empty (value INTEGER)')
                connection.commit()
            finally:
                connection.close()
            query = 'WITH RECURSIVE seq(x) AS (SELECT 1 UNION ALL SELECT x + 1 FROM seq WHERE x < 1001) SELECT x FROM seq'
            with self.assertRaisesRegex(ValueError, '1,000 data rows'):
                api.run_sqlite_select('rows', query)

    def test_authenticated_import_validates_https_and_credentials_before_network(self):
        api = DesktopAPI()
        for url in ('http://example.test/data.csv', 'https://user:pass@example.test/data.csv', 'https://example.test/data.csv#fragment'):
            with self.subTest(url=url):
                with self.assertRaisesRegex(ValueError, 'HTTPS'):
                    api.fetch_authenticated_data(url, 'csv', 'none')
        with self.assertRaisesRegex(ValueError, 'username and password'):
            api.fetch_authenticated_data('https://example.test/data.csv', 'csv', 'basic', '', 'secret')
        with self.assertRaisesRegex(ValueError, 'line breaks'):
            api.fetch_authenticated_data('https://example.test/data.csv', 'csv', 'bearer', '', 'token\r\ninjected')
        with self.assertRaisesRegex(ValueError, 'Bearer authentication'):
            api.fetch_authenticated_data('https://example.test/data.csv', 'csv', 'bearer', '', 'bad token')

    def test_filename_allowlist(self):
        self.assertEqual(safe_suggested_name('Sales / Q4.xlsx'), 'Sales _ Q4.xlsx')
        self.assertEqual(safe_suggested_name('book.xlsm'), 'book.xlsm')
        self.assertEqual(safe_suggested_name('chart.svg'), 'chart.svg')
        self.assertRaises(ValueError, safe_suggested_name, 'script.exe')
        self.assertRaises(ValueError, safe_suggested_name, '../secret.txt')

    def test_precompiled_assets(self):
        validate_dist(ROOT / 'dist')
        with tempfile.TemporaryDirectory() as folder:
            with self.assertRaises(FileNotFoundError):
                validate_dist(Path(folder))

    def test_loopback_and_content_type_no_directory_listing(self):
        server = create_local_server(ROOT / 'dist', 0)
        self.assertEqual(server.server_address[0], '127.0.0.1')
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        try:
            conn = HTTPConnection('127.0.0.1', server.server_port, timeout=3)
            conn.request('GET', '/main.js')
            response = conn.getresponse()
            self.assertEqual(response.status, 200)
            self.assertIn('javascript', response.getheader('Content-Type', ''))
            self.assertEqual(response.getheader('X-Content-Type-Options'), 'nosniff')
            response.read()
            conn.request('GET', '/engine/')
            resp = conn.getresponse()
            self.assertEqual(resp.status, 404)
            resp.read()
            conn.close()
        finally:
            server.shutdown(); server.server_close(); thread.join(timeout=3)

    def test_save_dialog_atomic_and_cancel(self):
        class FakeWindow:
            def __init__(self, choice):
                self.choice = choice
                self.calls = 0
            def create_file_dialog(self, *a, **kw):
                self.calls += 1
                return self.choice
        fake_module = types.SimpleNamespace(FileDialog=types.SimpleNamespace(SAVE=123))
        original = sys.modules.get('webview')
        sys.modules['webview'] = fake_module
        try:
            with tempfile.TemporaryDirectory() as folder:
                path = Path(folder) / 'test.xlsx'
                fake = FakeWindow((str(path),))
                desktop_module._desktop_window = fake
                api = DesktopAPI()
                self.assertFalse(hasattr(api, 'window'))
                self.assertEqual(api.save_file('test.xlsx', 'application/zip', base64.b64encode(b'PK-test').decode()), 'saved')
                self.assertEqual(path.read_bytes(), b'PK-test')
                self.assertEqual(fake.calls, 1)
                self.assertRaises(ValueError, api.save_file, 'bad.exe', 'text/plain', '')
                self.assertRaises(ValueError, api.save_file, 'chart.svg', 'image/svg+xml', base64.b64encode(b'<script>alert(1)</script>').decode())
                self.assertRaises(ValueError, api.save_file, 'chart.svg', 'image/svg+xml', base64.b64encode(b'<svg onload="alert(1)" />').decode())
                self.assertRaises(ValueError, api.save_file, 'test.xlsx', 'application/zip', base64.b64encode(b'not-zip').decode())
                desktop_module._desktop_window = FakeWindow(None)
                self.assertEqual(DesktopAPI().save_file('test.xlsx', '', base64.b64encode(b'PK-safe').decode()), 'cancelled')
        finally:
            desktop_module._desktop_window = None
            if original is None: sys.modules.pop('webview', None)
            else: sys.modules['webview'] = original

    def test_desktop_workflow_declared(self):
        src = (ROOT / 'src/main.ts').read_text(encoding='utf-8')
        self.assertIn('pywebview?.api?.save_file', src)
        self.assertIn('native.save_file(name, type, btoa(encoded))', src)
        wf = (ROOT / '.github/workflows/windows-desktop-release.yml').read_text(encoding='utf-8')
        self.assertIn('runs-on: windows-latest', wf)
        self.assertIn('GridForge-Portable-win64.zip', wf)
        self.assertIn('SHA256SUMS.txt', wf)
        iss = (ROOT / 'desktop/installer/GridForge.iss').read_text(encoding='utf-8')
        self.assertIn('PrivilegesRequired=lowest', iss)
        self.assertIn('Uninstall GridForge', iss)
        self.assertIn('GridForgeExcelBridge.exe', iss)
        self.assertEqual(PORT, 0)
        # The packaged app binds an ephemeral port so multiple launches and
        # stale instances cannot contend for one fixed local port.
        server = create_local_server(ROOT / 'dist')
        try:
            self.assertEqual(server.server_address[0], '127.0.0.1')
            self.assertGreater(server.server_port, 0)
        finally:
            server.server_close()


if __name__ == '__main__':
    unittest.main()
