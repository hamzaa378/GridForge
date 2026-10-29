#!/usr/bin/env python3
"""Chromium check of desktop's JS-to-native save handoff (no file dialog needed)."""
from __future__ import annotations
import base64, json, os, pathlib, re
from playwright.sync_api import sync_playwright
ROOT=pathlib.Path(__file__).resolve().parent.parent
DIST=ROOT/'dist'

with sync_playwright() as p:
    opts={'headless':True, 'args':['--no-sandbox']}
    if os.environ.get('GRIDFORGE_CHROMIUM'):
        opts['executable_path']=os.environ['GRIDFORGE_CHROMIUM']
    elif pathlib.Path('/usr/bin/chromium').exists():
        opts['executable_path']='/usr/bin/chromium'
    browser=p.chromium.launch(**opts)
    try:
        page=browser.new_page(viewport={'width':1280,'height':800})
        errors=[]
        page.on('pageerror',lambda e:errors.append(str(e)))
        imports={}
        for module in DIST.rglob('*.js'):
            source=module.read_text(encoding='utf-8')
            def repl(m):
                target=(module.parent/m.group(2)).resolve().relative_to(DIST.resolve()).as_posix()
                return m.group(1)+'gridforge/'+target+m.group(3)
            source=re.sub(r"(from\s+['\"])(\.[^'\"]+)(['\"])",repl,source)
            imports['gridforge/'+module.relative_to(DIST).as_posix()]='data:text/javascript;base64,'+base64.b64encode(source.encode()).decode()
        html=(DIST/'index.html').read_text(encoding='utf-8')
        html=html.replace('<link rel="stylesheet" href="./styles.css"/>','<style>'+(DIST/'styles.css').read_text(encoding='utf-8')+'</style>')
        html=html.replace('<script type="module" src="./main.js"></script>','')
        injection='''<script>
          window.desktopSaveCalls=[];
          window.pywebview={api:{save_file: async (name,type,encoded)=>{
            window.desktopSaveCalls.push({name,type,encoded});return 'saved';
          }}};
        </script>'''
        html=html.replace('</body>',injection+f'<script type="importmap">{json.dumps({"imports":imports})}</script><script type="module">import "gridforge/main.js";</script></body>')
        page.set_content(html,wait_until='load',timeout=20000)
        page.locator('[data-ref="D4"]').wait_for(timeout=10000)
        page.locator('[data-ribbon="File"]').click()
        page.locator('#btn-xlsx').click()
        page.wait_for_function('window.desktopSaveCalls.length === 1',timeout=5000)
        payload=page.evaluate('window.desktopSaveCalls[0]')
        assert payload['name'].endswith('.xlsx'),payload['name']
        assert base64.b64decode(payload['encoded']).startswith(b'PK'),'Downloaded XLSX bytes missing ZIP header'
        assert not errors, errors
        print('PASS: built app routes XLSX Save to desktop API with valid Base64 Office ZIP payload')
    finally:
        browser.close()
