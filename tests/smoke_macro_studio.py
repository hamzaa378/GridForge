"""Windows/Linux browser smoke test. Requires pip install playwright + playwright install chromium.

Uses the real local Python launcher, so it also exercises its --macro-workbench URL.
Some hosted browser environments prohibit localhost navigation; that is an environment block.
"""
from __future__ import annotations
import socket
import subprocess
import sys
from pathlib import Path
from playwright.sync_api import sync_playwright
ROOT=Path(__file__).resolve().parent.parent
with socket.socket() as s:
    s.bind(('127.0.0.1',0))
    PORT=s.getsockname()[1]
server=subprocess.Popen([sys.executable,'run_gridforge.py','--port',str(PORT),'--no-browser','--macro-workbench'],cwd=ROOT,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
try:
    with sync_playwright() as p:
        try:browser=p.chromium.launch(headless=True)
        except Exception:
            if sys.platform=='win32':raise
            browser=p.chromium.launch(executable_path='/usr/bin/chromium',headless=True,args=['--no-sandbox'])
        page=browser.new_page()
        errors=[]
        page.on('pageerror',lambda exc:errors.append(str(exc)))
        # Server startup can take a moment; navigation policy must permit loopback.
        page.goto(f'http://127.0.0.1:{PORT}/?macro-workbench=1',wait_until='load',timeout=25000)
        page.wait_for_timeout(400)
        assert page.locator('#macro-dialog').evaluate('(el)=>el.open'),'Macro Studio dialog did not open'
        assert page.locator('#macro-template option').count()>=6
        page.locator('#macro-template').select_option('parameter')
        page.locator('#macro-add-template').click()
        assert page.locator('#macro-select option').count()>=1
        page.locator('#macro-params').fill('{"tax":"0.15"}')
        page.locator('#macro-preview').click()
        assert page.locator('#macro-state').inner_text().startswith('Preview:')
        page.locator('#macro-run').click()
        assert not errors,errors
        print('PASS Macro Studio startup, template, parameterized preview and execution; no browser exceptions')
        browser.close()
finally:
    server.terminate()
    try:server.wait(timeout=5)
    except subprocess.TimeoutExpired:server.kill()
