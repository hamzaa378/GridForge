"""Regression tests runnable on Linux and Windows without a graphical desktop."""
import base64
from functools import partial
from http.client import HTTPConnection
from pathlib import Path
import re
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
        self.assertIn('interface GridForgeDesktopAPI', src)
        self.assertIn('native.save_file(name,type,btoa(encoded))', src)
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
