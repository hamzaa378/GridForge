#!/usr/bin/env python3
"""No Excel installation required: test bridge auth, CORS and XLSM validation."""
import base64
import http.client
import importlib.util
import io
import json
import socket
from pathlib import Path
from threading import Thread
from zipfile import ZipFile

SOURCE = Path(__file__).with_name('excel_vba_bridge.py')
spec = importlib.util.spec_from_file_location('excel_vba_bridge', SOURCE)
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
with socket.socket() as conn:
    conn.bind(('127.0.0.1', 0))
    module.PORT = conn.getsockname()[1]

server = module.HTTPServer((module.HOST, module.PORT), module.Handler)
worker = Thread(target=server.serve_forever, daemon=True)
worker.start()


def request(method, path, payload=None, *, token='', origin='http://127.0.0.1:4174'):
    client = http.client.HTTPConnection(module.HOST, module.PORT, timeout=3)
    headers = {'Origin': origin, 'Host': f'127.0.0.1:{module.PORT}', 'Content-Type': 'application/json', 'X-GridForge-Token': token}
    raw = json.dumps(payload).encode() if payload is not None else None
    client.request(method, path, raw, headers)
    response = client.getresponse()
    data = response.read()
    client.close()
    return response.status, json.loads(data) if data else {}

try:
    assert request('GET', '/health')[0] == 200
    status, _ = request('POST', '/run', {}, token='bad-token')
    assert status == 403
    status, _ = request('POST', '/run', {}, token=module.TOKEN, origin='https://malicious.example')
    assert status == 403
    status, _ = request('POST', '/run', {}, token=module.TOKEN)
    assert status == 400
    status, _ = request('POST', '/run', {'macro': 'Bad;del', 'xlsm_base64': 'AA=='}, token=module.TOKEN)
    assert status == 400
    assert module.MACRO_RE.fullmatch('Module1.GenerateReport')
    assert not module.MACRO_RE.fullmatch('Module1.GenerateReport(123)')
    with io.BytesIO() as out:
        with ZipFile(out, 'w') as archive:
            archive.writestr('[Content_Types].xml', 'application/vnd.ms-excel.sheet.macroEnabled.main+xml')
            archive.writestr('xl/workbook.xml', '<workbook/>')
        module.validate_xlsm(out.getvalue())
    try:
        module.validate_xlsm(b'not an xlsm')
        raise AssertionError('Malformed XLSM accepted')
    except ValueError:
        pass
    print('PASS VBA bridge authentication, origin restriction, macro-name validation, and XLSM preflight (8 checks)')
finally:
    server.shutdown()
    server.server_close()
    worker.join(timeout=5)
