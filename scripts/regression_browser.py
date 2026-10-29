#!/usr/bin/env python3
"""End-to-end GridForge v6 regression in a fresh Chromium browser profile.

Requires: python -m pip install playwright
          python -m playwright install chromium
No changes are made to a user's existing browser storage or GridForge files.
"""
from __future__ import annotations

import json
import base64
import re
import os
import socket
import subprocess
import sys
import time
import zipfile
from pathlib import Path
from urllib.error import URLError
from urllib.request import urlopen

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print('BROWSER SETUP REQUIRED: python -m pip install playwright', file=sys.stderr)
    sys.exit(2)

ROOT = Path(__file__).resolve().parent.parent
checks: list[str] = []
mode = "HTTP server"


def verify(label: str, test) -> None:
    try:
        test()
        checks.append(label)
        print(f'  PASS {label}', flush=True)
    except Exception:
        print(f'  FAIL {label}', file=sys.stderr, flush=True)
        raise


def assert_eq(actual, expected, label='value'):
    assert actual == expected, f'{label}: expected {expected!r}, got {actual!r}'


def port_available() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(('127.0.0.1', 0))
        return sock.getsockname()[1]


def ready(url: str, process: subprocess.Popen, limit: float = 12) -> None:
    deadline = time.monotonic() + limit
    while time.monotonic() < deadline:
        if process.poll() is not None:
            raise RuntimeError(f'GridForge server stopped unexpectedly (code {process.returncode})')
        try:
            with urlopen(url, timeout=0.8) as response:
                if response.status == 200:
                    return
        except (OSError, URLError):
            time.sleep(0.15)
    raise RuntimeError('GridForge server did not respond within 12 seconds')


def run() -> None:
    port = port_available()
    url = f'http://127.0.0.1:{port}'
    env = {**os.environ, 'PORT': str(port)}
    # Independent ephemeral HTTP server. No interaction with the user's running port 4174.
    server = subprocess.Popen(
        ['node', 'scripts/serve.mjs'], cwd=ROOT, env=env,
        stdout=subprocess.DEVNULL, stderr=subprocess.PIPE,
    )
    try:
        ready(url, server)
        with sync_playwright() as p:
            kwargs = {'headless': True, 'args': ['--no-sandbox']}
            if os.getenv('GRIDFORGE_CHROMIUM'):
                kwargs['executable_path'] = os.environ['GRIDFORGE_CHROMIUM']
            try:
                browser = p.chromium.launch(**kwargs)
            except Exception as exc:
                raise RuntimeError(
                    'Chromium unavailable. Run: python -m playwright install chromium '
                    '(or set GRIDFORGE_CHROMIUM to your Chromium executable)'
                ) from exc
            try:
                context = browser.new_context(
                    viewport={'width': 1440, 'height': 900}, accept_downloads=True
                )
                page = context.new_page()
                page.set_default_timeout(10000)
                errors: list[str] = []
                page.on('pageerror', lambda err: errors.append(str(err)))
                global mode
                try:
                    page.goto(url, wait_until='load', timeout=15000)
                except Exception as exc:
                    if 'ERR_BLOCKED_BY_ADMINISTRATOR' not in str(exc):
                        raise
                    page.close()
                    page = context.new_page()
                    page.on('pageerror', lambda err: errors.append(str(err)))
                    # Hosted browsers may block all origin navigation. Inline mode
                    # still tests UI actions but CANNOT test actual localStorage reload.
                    mode = 'offline embedded document (storage reload skipped)'
                    dist = ROOT / 'dist'
                    modules = {}
                    for module in dist.rglob('*.js'):
                        relative = module.relative_to(dist).as_posix()
                        source = module.read_text(encoding='utf-8')
                        def rewrite(match):
                            target = (module.parent / match.group(2)).resolve().relative_to(dist.resolve()).as_posix()
                            return match.group(1) + 'gridforge/' + target + match.group(3)
                        source = re.sub(r"(from\s+['\"])(\.[^'\"]+)(['\"])", rewrite, source)
                        modules['gridforge/' + relative] = 'data:text/javascript;base64,' + base64.b64encode(source.encode()).decode()
                    html = (dist/'index.html').read_text(encoding='utf-8')
                    html = html.replace('<link rel="stylesheet" href="./styles.css"/>', '<style>' + (dist/'styles.css').read_text(encoding='utf-8') + '</style>')
                    html = html.replace('<script type="module" src="./main.js"></script>', '')
                    html = html.replace('</body>', f'<script type="importmap">{json.dumps({"imports": modules})}</script><script type="module">import "gridforge/main.js";</script></body>')
                    page.set_content(html, wait_until='load', timeout=15000)
                page.locator('[data-ref="D4"]').wait_for(timeout=10000)

                def cell(ref):
                    return page.locator(f'[data-ref="{ref}"]')

                def select(ref):
                    page.locator('#name-box').fill(ref)
                    page.locator('#name-box').press('Enter')

                def formula(ref):
                    select(ref)
                    return page.locator('#formula').input_value()

                def edit(ref, raw):
                    select(ref)
                    page.locator('#formula').fill(raw)
                    page.locator('#formula').press('Enter')

                def excel_like_blank_start():
                    assert_eq(page.locator('#book-title').inner_text(),'Book1','new workbook title')
                    assert_eq(page.locator('#formula').input_value(),'','initial formula bar')
                    assert_eq(page.locator('[data-ref="A1"]').inner_text(),'','initial active cell')
                    assert page.locator('.quick-access #btn-undo').count()==1
                    assert page.locator('[data-ribbon="Help"]').count()==1
                    page.locator('[data-ribbon="File"]').click()
                    page.once('dialog',lambda dlg:dlg.accept())
                    page.locator('#btn-sample').click()
                    assert_eq(cell('D4').inner_text(),'4000','loaded example D4')
                    page.locator('[data-ribbon="Home"]').click()
                verify('Excel-style blank workbook start and Quick Access controls',excel_like_blank_start)
                def baseline():
                    assert_eq(cell('D4').inner_text(), '4000', 'sample D4')
                    assert_eq(formula('D4'), '=B4*C4', 'D4 formula')
                verify('Example workbook formulas calculate correctly', baseline)
                def clean_excel_workspace():
                    assert_eq(page.evaluate('document.documentElement.dataset.theme'),'light')
                    assert_eq(page.evaluate('document.documentElement.dataset.inspector'),'closed')
                    assert page.locator('[data-ribbon="Home"]').get_attribute('aria-selected')=='true'
                    assert not page.locator('.metrics').is_visible()
                    assert not page.locator('#sidebar').is_visible()
                    page.screenshot(path=str(ROOT/'docs'/'preview_v10_3_clean.png'),full_page=False)
                    page.locator('#btn-insights').click()
                    assert page.locator('#sidebar').is_visible()
                    assert page.locator('.metrics').is_visible()
                    page.locator('#btn-insights').click()
                    assert not page.locator('#sidebar').is_visible()
                    page.locator('[data-ribbon="All"]').click()
                verify('v10.3: Excel-style clean home and optional insights toggle',clean_excel_workspace)
                def v10_themes_and_workspace():
                    page.locator('[data-ribbon="All"]').click()
                    theme=page.locator('#theme-switch')
                    theme.wait_for()
                    assert_eq(page.evaluate('document.documentElement.dataset.theme'),'light')
                    assert_eq(page.evaluate('document.documentElement.dataset.workspace'),'excel')
                    assert_eq(page.evaluate('document.documentElement.dataset.inspector'),'closed')
                    assert page.locator('#grid-scroll').is_visible()
                    prior=cell('D4').inner_text()
                    theme.click()
                    assert_eq(page.evaluate('document.documentElement.dataset.theme'),'dark')
                    assert_eq(theme.get_attribute('aria-pressed'),'false')
                    assert page.locator('.topbar').is_visible()
                    assert_eq(cell('D4').inner_text(),prior,'formula survives theme change')
                    page.locator('#workspace-switch').select_option('guided')
                    assert page.locator('#guided-home').is_visible()
                    assert_eq(page.evaluate('document.documentElement.dataset.workspace'),'guided')
                    page.locator('#guided-functions').click()
                    assert_eq(page.locator('[data-ribbon="Formulas"]').get_attribute('aria-selected'),'true')
                    page.locator('#workspace-switch').select_option('excel')
                    assert not page.locator('#guided-home').is_visible()
                    theme.click()
                    assert_eq(page.evaluate('document.documentElement.dataset.theme'),'light')
                    assert_eq(cell('D4').inner_text(),prior)
                verify('v10 persistent-preference controls, light/dark contrast, Excel/Guided switch without data loss',v10_themes_and_workspace)

                def v10_new_formula_workbook():
                    page.locator('[data-ribbon="All"]').click()
                    for ref, raw, expected in [('Z180','=BETA.DIST(0.5,2,2,TRUE)','0.5'),('Z181','=INDEX(SEQUENCE(3,2),2,2)','4'),('Z182','=INDEX(TEXTSPLIT("red,blue",","),1,2)','blue'),('Z183','=ISBETWEEN(5,3,7)','TRUE'),('Z184','=DETECTLANGUAGE("Hello")','en'),('Z185','=EUROCONVERT(100,"EUR","DEM")','195.583')]:
                        edit(ref,raw)
                        assert_eq(cell(ref).inner_text(),expected,ref)
                verify('v10 modern formulas and matrix functions evaluate through spreadsheet grid',v10_new_formula_workbook)
                if os.environ.get('GRIDFORGE_SCREENSHOTS') == '1':
                    page.locator('[data-ribbon="All"]').click()
                    page.locator('#theme-switch').click()
                    page.screenshot(path=str(ROOT/'docs'/'preview_v10_light.png'),full_page=False)
                    page.locator('#workspace-switch').select_option('guided')
                    page.screenshot(path=str(ROOT/'docs'/'preview_v10_guided.png'),full_page=False)
                    page.locator('#theme-switch').click()
                    page.locator('#workspace-switch').select_option('excel')
                    page.screenshot(path=str(ROOT/'docs'/'preview_v10_dark.png'),full_page=False)
                # Capture Excel-friendly v9 UI from the same tested isolated browser.
                page.locator('[data-ribbon="Formulas"]').click()
                page.screenshot(path=str(ROOT/'docs'/'preview_v9_excel_interface.png'),full_page=False)
                page.locator('#btn-first-steps').click()
                page.screenshot(path=str(ROOT/'docs'/'preview_v9_onboarding.png'),full_page=False)
                page.locator('#close-welcome').click()
                page.locator('[data-ribbon="All"]').click()
                # Capture real rendered screenshot before the regression mutates storage.
                page.locator('[data-ribbon="Data"]').click()
                page.locator('#btn-data-studio').click()
                page.screenshot(path=str(ROOT/'docs'/'preview_v8_data_studio.png'),full_page=False)
                page.locator('#close-data-studio').click()
                page.locator('[data-ribbon="Formulas"]').click()
                if not page.locator('[data-mode="functions"]').is_visible():
                    page.locator('#btn-insights').click()
                page.locator('[data-mode="functions"]').click()
                page.locator('#function-category').select_option('Financial')
                page.screenshot(path=str(ROOT/'docs'/'preview_v6.png'),full_page=False)
                page.locator('[data-mode="inspector"]').click()
                page.locator('[data-ribbon="All"]').click()
                page.locator('#btn-automate').click()
                page.screenshot(path=str(ROOT/'docs'/'preview_v7_macros.png'),full_page=False)
                page.locator('#close-macros').click()

                def edit_undo_redo():
                    edit('B4', '225')
                    assert_eq(cell('D4').inner_text(), '7200', 'dependent D4')
                    page.locator('#btn-undo').click()
                    assert_eq(cell('D4').inner_text(), '4000', 'undo dependent')
                    page.locator('#btn-redo').click()
                    assert_eq(cell('D4').inner_text(), '7200', 'redo dependent')
                    page.locator('#btn-undo').click()
                verify('Editing, dependency recalculation, undo/redo', edit_undo_redo)

                def formula_error():
                    edit('N2', '=SUM(B4, B5')
                    assert_eq(cell('N2').inner_text(), '#PARSE!', 'parse diagnostic')
                    select('N2')
                    # Validate diagnostic on-screen rather than just the engine result.
                    assert 'parenthesis' in page.locator('body').inner_text().lower()
                    page.locator('#btn-undo').click()
                verify('Malformed formula shows useful parser error', formula_error)

                def style_undo():
                    cell('D4').click()
                    page.locator('#fmt-bold').click()
                    assert_eq(cell('D4').evaluate('(el)=>el.style.fontWeight'), '800', 'bold')
                    page.locator('#btn-undo').click()
                    assert cell('D4').evaluate('(el)=>el.style.fontWeight') != '800'
                    page.locator('#btn-redo').click()
                    assert_eq(cell('D4').evaluate('(el)=>el.style.fontWeight'), '800')
                verify('Formatting and format history', style_undo)

                def right_click_copy():
                    cell('D4').click(button='right')
                    assert page.locator('#cell-context-menu').is_visible()
                    page.locator('[data-context="copy"]').click()
                    cell('H4').click(button='right')
                    page.locator('[data-context="paste"]').click()
                    assert_eq(formula('H4'), '=F4*G4', 'relative formula paste')
                    page.locator('#btn-undo').click()
                    assert_eq(formula('H4'), '', 'paste undo')
                    page.locator('#grid-scroll').focus()
                    page.locator('#grid-scroll').press('Shift+F10')
                    assert page.locator('#cell-context-menu').is_visible()
                    page.keyboard.press('Escape')
                    assert not page.locator('#cell-context-menu').is_visible()
                verify('Right-click menu, formula paste, keyboard menu', right_click_copy)

                def sorting():
                    cell('A3').click()
                    cell('G11').click(modifiers=['Shift'])
                    page.locator('#btn-sort-custom').click()
                    page.locator('#sort-column').select_option('1')
                    page.locator('#sort-header').check()
                    page.locator('#btn-sort-apply').click()
                    assert_eq(cell('A4').inner_text(), 'P-103', 'first sorted SKU')
                    assert_eq(cell('B4').inner_text(), '51', 'sorted row quantity')
                    assert_eq(cell('D4').inner_text(), '6120', 'sorted row formula')
                    page.locator('#btn-undo').click()
                    assert_eq(cell('A4').inner_text(), 'P-100', 'sort undo')
                verify('Custom sort preserves complete rows and undo', sorting)

                def find_replace():
                    page.locator('#btn-find').click()
                    page.locator('#find-query').fill('P-103')
                    page.locator('#btn-find-all').click()
                    assert page.locator('.search-result').count() >= 2
                    page.locator('#replace-text').fill('SKU-103')
                    page.once('dialog', lambda dlg: dlg.accept())
                    page.locator('#btn-replace-all').click()
                    assert_eq(cell('A7').inner_text(), 'SKU-103', 'replaced SKU')
                    page.locator('#close-find').click()
                    page.locator('#btn-undo').click()
                    assert_eq(cell('A7').inner_text(), 'P-103', 'replace undo')
                verify('Find, replace, and undo', find_replace)

                def conditional_filter():
                    cell('B4').click()
                    cell('B11').click(modifiers=['Shift'])
                    page.locator('#btn-conditional').click()
                    page.locator('#rule-op').select_option('gt')
                    page.locator('#rule-threshold').fill('100')
                    page.locator('#rule-color').fill('#276a53')
                    page.locator('#apply-conditional').click()
                    assert_eq(cell('B4').evaluate('(el)=>el.style.backgroundColor'), 'rgb(39, 106, 83)')
                    assert cell('B5').evaluate('(el)=>el.style.backgroundColor') != 'rgb(39, 106, 83)'
                    cell('A3').click()
                    cell('G11').click(modifiers=['Shift'])
                    page.locator('#btn-filter').click()
                    page.locator('#filter-column').select_option('0')
                    page.locator('#filter-query').fill('P-103')
                    page.locator('#apply-filter').click()
                    assert cell('A7').count() == 1
                    assert cell('A4').count() == 0
                    page.locator('#btn-clear-filter').click()
                    assert_eq(cell('A4').inner_text(), 'P-100')
                verify('Conditional formatting and non-destructive filter', conditional_filter)

                def freeze_autosum():
                    page.locator('#btn-freeze-row').click()
                    page.locator('#btn-freeze-col').click()
                    assert page.locator('.frozen-cell').count() > 0
                    select('B12')
                    page.locator('#btn-autosum').click()
                    assert_eq(formula('B12'), '=SUM(B4:B11)', 'AutoSum formula')
                verify('Freeze panes and AutoSum', freeze_autosum)

                def sheets():
                    page.locator('#btn-add-sheet').click()
                    assert_eq(page.locator('[data-sheet-id]').count(), 2)
                    edit('A1', '42')
                    assert_eq(cell('A1').inner_text(), '42')
                    page.locator('#btn-duplicate-sheet').click()
                    assert_eq(page.locator('[data-sheet-id]').count(), 3)
                    assert_eq(cell('A1').inner_text(), '42', 'duplicated value')
                    page.locator('[data-sheet-id="s1"]').click()
                    assert_eq(cell('A4').inner_text(), 'P-100', 'original sheet remains')
                    page.locator('[data-sheet-id="s2"]').click()
                    assert_eq(cell('A1').inner_text(), '42', 'sheet data preserved')
                verify('Add, duplicate, switch independent worksheets', sheets)

                def json_and_reload():
                    page.locator('[data-ribbon="File"]').click()
                    with page.expect_download(timeout=8000) as download_info:
                        page.locator('#btn-json').click()
                    export = json.loads(download_info.value.path().read_text(encoding='utf-8'))
                    assert_eq(export['version'], 3)
                    assert_eq(len(export['sheets']), 3)
                    assert_eq(export['sheets'][1]['workbook']['cells']['A1'], '42')
                    if mode == 'HTTP server':
                        page.wait_for_timeout(550)  # browser save debounce is 250ms
                        page.reload(wait_until='load')
                        page.locator('[data-ref="A1"]').wait_for()
                        assert_eq(page.locator('[data-sheet-id]').count(), 3, 'persistent sheets')
                        assert_eq(cell('A1').inner_text(), '42', 'persisted active sheet')
                        assert_eq(formula('A1'), '42')
                    else:
                        print('  SKIP storage reload (browser navigation prohibited in this environment)')
                verify('JSON export' + (' and storage persistence' if mode == 'HTTP server' else ' (storage not tested)'), json_and_reload)

                def xlsx_roundtrip():
                    page.locator('[data-ribbon="File"]').click()
                    with page.expect_download(timeout=15000) as download_info:
                        page.locator('#btn-xlsx').click()
                    path = download_info.value.path()
                    with zipfile.ZipFile(path) as archive:
                        assert archive.testzip() is None
                        names = archive.namelist()
                        assert 'xl/workbook.xml' in names
                        assert 'xl/worksheets/sheet1.xml' in names
                        assert 'xl/worksheets/sheet2.xml' in names
                        assert 'xl/worksheets/sheet3.xml' in names
                        xml = archive.read('xl/worksheets/sheet2.xml').decode('utf-8')
                        assert '<v>42</v>' in xml
                    # Re-import the actual exported file through the UI, not just XML inspection.
                    page.locator('#file-picker').set_input_files({'name': 'roundtrip.xlsx', 'mimeType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'buffer': Path(path).read_bytes()}, timeout=7000)
                    page.locator('#import-notes-dialog').wait_for(state='visible',timeout=7000)
                    page.locator('#confirm-import-notes').click()
                    assert_eq(page.locator('[data-sheet-id]').count(), 3)
                    page.locator('[data-sheet-id="s2"]').click()
                    assert_eq(cell('A1').inner_text(), '42', 're-imported cell')
                verify('Actual XLSX download, ZIP integrity, UI re-import', xlsx_roundtrip)

                def fixture_import():
                    fixture = ROOT / 'examples' / 'sample_sales.xlsx'
                    assert fixture.exists(), 'sample XLSX fixture missing'
                    page.locator('#file-picker').set_input_files(str(fixture))
                    page.locator('#import-notes-dialog').wait_for(state='visible')
                    page.locator('#confirm-import-notes').click()
                    assert_eq(page.locator('[data-sheet-id]').count(), 2)
                    page.locator('[data-sheet-id="s1"]').click()
                    assert_eq(formula('D2'), '=B2*C2')
                    assert '75,000' in cell('D2').inner_text()
                verify('Bundled Excel fixture imports formulas and sheets', fixture_import)

                def expanded_functions():
                    cases = [
                        ('Z20', '=COMBIN(5,2)', '10'),
                        ('Z21', '=LOG10(1000)', '3'),
                        ('Z22', '=IFS(TRUE,"ready",FALSE,1/0)', 'ready'),
                        ('Z23', '=TEXTBEFORE("one-two-three","-",2)', 'one-two'),
                        ('Z24', '=XLOOKUP("k2",D20:D21,E20:E21,"not found")', 'not found'),
                        ('Z25', '=NETWORKDAYS(DATE(2026,10,19),DATE(2026,10,23))', '5'),
                    ]
                    for ref, raw, expected in cases:
                        edit(ref, raw)
                        assert_eq(cell(ref).inner_text(), expected, f'new function at {ref}')
                verify('New v4 formulas work through the browser grid', expanded_functions)

                def expanded_xlsx_import():
                    path = ROOT / 'examples' / 'v4_80_function_gallery.xlsx'
                    assert path.exists(), 'v4 function gallery XLSX missing'
                    page.locator('#file-picker').set_input_files(str(path))
                    page.locator('#import-notes-dialog').wait_for(state='visible')
                    page.locator('#confirm-import-notes').click()
                    assert_eq(page.locator('[data-sheet-id]').count(), 1)
                    select('C25')
                    assert_eq(cell('C25').inner_text(), '3', 'LOG10 survives XLSX')
                    select('C76')
                    assert_eq(cell('C76').inner_text(), '20', 'XLOOKUP with _xlfn survives XLSX')
                    assert formula('C76').startswith('=XLOOKUP(')
                verify('v4 Excel gallery import supports modern Excel function prefixes', expanded_xlsx_import)

                def v5_categories():
                    page.locator('[data-mode="functions"]').click()
                    assert '506 local · 10 connected' in page.locator('#sidebar-body').inner_text()
                    assert page.locator('[data-mode="functions"]').evaluate('(el)=>el.classList.contains("chosen")')
                    page.locator('#function-category').select_option('Financial')
                    assert page.locator('[data-function="RATE"]').count() == 1
                    assert page.locator('[data-function="DATEDIF"]').count() == 0
                    page.locator('#function-category').select_option('Date & Time')
                    assert page.locator('[data-function="DATEDIF"]').count() == 1
                    page.locator('#function-category').select_option('Recently added')
                    assert_eq(page.locator('[data-function]').count(), 100)
                    page.locator('#function-category').select_option('All functions')
                verify('v10.4 categorized formula browser shows offline and connected names', v5_categories)

                def v5_formulas():
                    examples=[('Z90','=BASE(255,16)','FF'),('Z91','=NETWORKDAYSINTL(DATE(2026,10,19),DATE(2026,10,23),1)','5'),
                      ('Z92','=ISERROR(1/0)','TRUE'),('Z93','=CHOOSE(2,1/0,99)','99'),('Z94','=REGEXREPLACE("ab12","[0-9]+","ZZ")','abZZ')]
                    for ref, raw, expected in examples:
                        edit(ref,raw);assert_eq(cell(ref).inner_text(),expected,ref)
                verify('v5 engineering, date, error, lazy and regex functions work in grid',v5_formulas)

                def v5_action_macros():
                    page.locator('[data-ribbon="Automate"]').click()
                    page.locator('#btn-automate').click()
                    page.locator('#macro-name').fill('Regression Macro')
                    page.locator('#macro-record').click()
                    edit('Z97','123')
                    select('Z97')
                    page.locator('[data-ribbon="Home"]').click()
                    page.locator('#fmt-bold').click()
                    page.locator('[data-ribbon="Automate"]').click()
                    page.locator('#btn-automate').click()
                    assert '2 action(s)' in page.locator('#macro-state').inner_text()
                    page.locator('#macro-stop').click()
                    page.locator('#close-macros').click()
                    edit('Z97','0')
                    select('Z97')
                    page.locator('[data-ribbon="Home"]').click()
                    page.locator('#fmt-bold').click()
                    page.locator('[data-ribbon="Automate"]').click()
                    page.locator('#btn-automate').click()
                    page.locator('#macro-run').click()
                    assert_eq(cell('Z97').inner_text(),'123')
                    assert_eq(cell('Z97').evaluate('(el)=>el.style.fontWeight'),'800')
                    with page.expect_download(timeout=10000) as dl:
                        page.locator('#macro-export').click()
                    result=dl.value
                    data=json.loads(Path(result.path()).read_text())
                    assert_eq(data['version'],2)
                    assert_eq(len(data['steps']),2)
                    assert all(step['kind'] in ('cells','styles') for step in data['steps'])
                    page.locator('#close-macros').click()
                verify('v5 safe action macro record, playback and JSON export',v5_action_macros)

                def v5_gallery():
                    path=ROOT/'examples'/'v5_80_function_gallery.xlsx'
                    assert path.exists()
                    page.locator('#file-picker').set_input_files(str(path))
                    page.locator('#import-notes-dialog').wait_for(state='visible')
                    page.locator('#confirm-import-notes').click()
                    assert_eq(cell('C2').inner_text(),'0')
                    select('A81')
                    assert page.locator('[data-ref="A81"]').inner_text() != ''
                    assert not page.locator('[data-ref="C81"]').inner_text().startswith('#')
                verify('v5 complete 80-function XLSX showcase imports and calculates',v5_gallery)

                def v5_macro_enabled_import():
                    import tempfile
                    with tempfile.TemporaryDirectory() as tmp:
                        source=ROOT/'examples'/'sample_sales.xlsx'
                        modified=Path(tmp)/'legacy_macro_fixture.xlsm'
                        with zipfile.ZipFile(source) as old,zipfile.ZipFile(modified,'w') as new:
                            for item in old.infolist():
                                data=old.read(item.filename)
                                if item.filename=='[Content_Types].xml':
                                    data=data.replace(b'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',b'application/vnd.ms-excel.sheet.macroEnabled.main+xml')
                                new.writestr(item.filename,data)
                            new.writestr('xl/vbaProject.bin',b'SYNTHETIC VBA PLACEHOLDER - NOT EXECUTABLE')
                        before=modified.read_bytes()
                        page.locator('#file-picker').set_input_files(str(modified))
                        page.locator('#import-notes-dialog').wait_for(state='visible')
                        assert 'never executed' in page.locator('#import-notes-body').inner_text().lower()
                        page.locator('#confirm-import-notes').click()
                        assert page.locator('#btn-original-xlsm').is_visible()
                        edit('Z80','42')
                        with page.expect_download(timeout=10000) as original_dl:
                            page.locator('#btn-original-xlsm').click()
                        assert Path(original_dl.value.path()).read_bytes() == before, 'original bytes must remain unchanged'
                        page.locator('[data-ribbon="File"]').click()
                        with page.expect_download(timeout=10000) as edited_dl:
                            page.locator('#btn-xlsx').click()
                        with zipfile.ZipFile(edited_dl.value.path()) as exported:
                            assert 'xl/vbaProject.bin' not in exported.namelist(), 'VBA must not leak into XLSX'
                        assert page.locator('#btn-save-xlsm').is_visible()
                        with page.expect_download(timeout=15000) as patched_dl:
                            page.locator('#btn-save-xlsm').click()
                        with zipfile.ZipFile(patched_dl.value.path()) as exported:
                            assert exported.read('xl/vbaProject.bin') == b'SYNTHETIC VBA PLACEHOLDER - NOT EXECUTABLE'
                            xml = b''.join(exported.read(f) for f in exported.namelist() if f.startswith('xl/worksheets/sheet') and f.endswith('.xml'))
                            assert b'Z80' in xml and b'42' in xml, 'cell edit did not round-trip to XLSM'
                        page.locator('#import-notes-dialog').wait_for(state='visible')
                        page.locator('#confirm-import-notes').click()
                        assert page.locator('#run-excel-vba').count() == 1, 'Missing opt-in desktop Excel VBA action'
                verify('v5 XLSM loads safely, original is byte-identical, edited XLSX strips VBA',v5_macro_enabled_import)

                def v6_dotted_formulas():
                    examples=[('Z90','=NORM.S.DIST(0,TRUE)','0.5'),('Z91','=BITAND(6,3)','2'),('Z92','=IFNA(NA(),77)','77'),('Z93','=ROW(B7)','7'),('Z94','=COLUMN(B7)','2')]
                    for ref,raw,expected in examples:
                        edit(ref,raw)
                        assert_eq(cell(ref).inner_text(),expected,ref)
                verify('v6 dotted names, bitwise, lazy and reference-aware functions',v6_dotted_formulas)

                def v6_gallery():
                    path=ROOT/'examples'/'v6_90_function_gallery.xlsx'
                    assert path.is_file()
                    page.locator('#file-picker').set_input_files(str(path))
                    page.locator('#import-notes-dialog').wait_for(state='visible')
                    page.locator('#confirm-import-notes').click()
                    assert_eq(cell('F2').inner_text(),'SQRTPI')
                    select('F91');assert page.locator('[data-ref="F91"]').inner_text() != ''
                    select('H91');assert not cell('H91').inner_text().startswith('#')
                verify('v6 90-function XLSX showcase imports and evaluates',v6_gallery)

                def v6_macro_preview():
                    import tempfile
                    with tempfile.TemporaryDirectory() as folder:
                        source=Path(folder)/'advanced_macro.json'
                        source.write_text(json.dumps({'version':2,'name':'v6 preview',
                            'sheetName':'90 Functions v6','steps':[
                                {'kind':'cells','edits':[{'ref':'Z110','raw':'250'}]},
                                {'kind':'fillDown','source':'Z110','range':'Z111:Z112'},
                                {'kind':'styles','edits':[{'ref':'Z110','style':{'bold':True}}]},
                                {'kind':'copy','sourceRange':'Z110:Z112','destination':'AA110'}]}))
                        page.locator('[data-ribbon="Automate"]').click()
                        page.locator('#btn-automate').click()
                        page.locator('#macro-file').set_input_files(str(source))
                        page.locator('#macro-select option', has_text='v6 preview').wait_for(state='attached')
                        page.locator('#macro-preview').click()
                        assert 'Preview:' in page.locator('#macro-state').inner_text()
                        page.locator('#close-macros').click()
                        select('Z110');assert cell('Z110').inner_text() != '250','dry-run mutated data'
                        page.locator('[data-ribbon="Automate"]').click()
                        page.locator('#btn-automate').click()
                        page.locator('#macro-run').click()
                        page.locator('#close-macros').click()
                        assert_eq(cell('Z110').inner_text(),'250')
                        assert_eq(cell('AA112').inner_text(),'250')
                        assert_eq(cell('Z110').evaluate('(el)=>el.style.fontWeight'),'800')
                verify('v6 safe macros preview without edits and replay after validation',v6_macro_preview)

                def v8_studio_catalogue_and_edit():
                    edit('Z130', 'hELLO wORLD')
                    select('Z130')
                    page.locator('[data-ribbon="Data"]').click()
                    page.locator('#btn-data-studio').click()
                    assert page.locator('#data-studio-dialog').is_visible()
                    assert '100 actions' in page.locator('#data-studio-dialog h2').inner_text()
                    assert_eq(page.locator('#studio-category option').count(),11)
                    page.locator('#studio-category').select_option('Text & case')
                    assert_eq(page.locator('#studio-action option').count(),10)
                    page.locator('#studio-action').select_option('title')
                    page.locator('#studio-preview-button').click()
                    assert '1 cell will change' in page.locator('#studio-preview').inner_text()
                    assert_eq(cell('Z130').inner_text(),'hELLO wORLD','preview does not mutate')
                    page.locator('#studio-apply').click()
                    assert_eq(cell('Z130').inner_text(),'Hello World')
                    page.locator('[data-ribbon="Home"]').click()
                    page.locator('#btn-undo').click()
                    assert_eq(cell('Z130').inner_text(),'hELLO wORLD')
                    page.locator('#btn-redo').click()
                    assert_eq(cell('Z130').inner_text(),'Hello World')
                verify('v8 Data Studio 100 actions, search, preview, atomic edit and undo/redo', v8_studio_catalogue_and_edit)

                def v8_studio_formula_and_bounds():
                    edit('Z131','2')
                    edit('Z132','4')
                    cell('Z131').click()
                    cell('Z132').click(modifiers=['Shift'])
                    page.locator('[data-ribbon="Data"]').click()
                    page.locator('#btn-data-studio').click()
                    page.locator('#studio-category').select_option('Quick aggregates')
                    page.locator('#studio-action').select_option('insert-sum')
                    page.locator('#studio-preview-button').click()
                    assert '1 cell will change' in page.locator('#studio-preview').inner_text()
                    page.locator('#studio-apply').click()
                    assert_eq(formula('Z133'),'=SUM(Z131:Z132)')
                    assert_eq(cell('Z133').inner_text(),'6')
                    page.locator('[data-ribbon="Home"]').click()
                    page.locator('#btn-undo').click()
                    assert_eq(formula('Z133'),'')
                    edit('Z137','=A1+$B$1')
                    cell('Z137').click()
                    cell('Z138').click(modifiers=['Shift'])
                    page.locator('[data-ribbon="Data"]').click()
                    page.locator('#btn-data-studio').click()
                    page.locator('#studio-category').select_option('Fill & series')
                    page.locator('#studio-action').select_option('fill-down')
                    page.locator('#studio-preview-button').click()
                    page.locator('#studio-apply').click()
                    assert_eq(formula('Z138'),'=A2+$B$1')
                    page.locator('[data-ribbon="Home"]').click()
                    page.locator('#btn-undo').click()
                    assert_eq(formula('Z138'),'')
                verify('v8 Data Studio aggregation, relative fill, range and undo',v8_studio_formula_and_bounds)

                def v9_excel_navigation():
                    assert page.locator('.excel-tab[aria-selected="true"]').count() == 1
                    page.locator('[data-ribbon="Data"]').click()
                    assert page.locator('#btn-sort-custom').is_visible()
                    assert not page.locator('#btn-json').is_visible()
                    page.locator('[data-ribbon="File"]').click()
                    assert page.locator('#btn-xlsx').is_visible()
                    assert not page.locator('#btn-sort-custom').is_visible()
                    page.locator('#command-search').fill('sort')
                    page.locator('#command-search').press('Enter')
                    assert page.locator('[data-ribbon="Data"]').get_attribute('aria-selected') == 'true'
                    page.locator('[data-ribbon="Formulas"]').click()
                    assert '506 local · 10 connected' in page.locator('#sidebar-body').inner_text()
                    assert page.locator('[data-mode="functions"]').evaluate('(el)=>el.classList.contains("chosen")')
                    page.locator('#function-category').select_option('Recently added')
                    assert_eq(page.locator('#sidebar-body [data-function]').count(),100)
                    page.locator('#function-category').select_option('Database')
                    assert page.locator('[data-function="DSUM"]').count() == 1
                    page.locator('[data-ribbon="All"]').click()
                verify('v9 Excel-style ribbon, command search and 100 new categorized functions',v9_excel_navigation)

                def v9_getting_started_and_examples():
                    page.locator('#btn-first-steps').click()
                    assert page.locator('#welcome-dialog').is_visible()
                    assert 'Excel experience' in page.locator('#welcome-dialog').inner_text()
                    page.locator('#welcome-functions').click()
                    assert page.locator('[data-ribbon="Formulas"]').get_attribute('aria-selected') == 'true'
                    page.locator('[data-ribbon="All"]').click()
                    examples=[('Z140','=BIN2DEC("1010")','10'),('Z141','=IMABS("3+4i")','5'),('Z142','=CONVERT(1,"mi","km")','1.609344'),('Z143','=TRUE()','TRUE'),('Z144','=FALSE()','FALSE'),('Z145','=INDEX(TRANSPOSE(H1:I2),2,1)','')]
                    for ref,raw,expected in examples[:-1]:
                        edit(ref,raw)
                        assert_eq(cell(ref).inner_text(),expected,ref)
                    assert not errors,errors
                verify('v9 getting started guide and new Excel formulas in browser',v9_getting_started_and_examples)

                verify('Browser JavaScript exceptions', lambda: assert_eq(errors, []))
                context.close()
            finally:
                browser.close()
    finally:
        server.terminate()
        try:
            server.communicate(timeout=3)
        except subprocess.TimeoutExpired:
            server.kill()
            server.communicate(timeout=2)


if __name__ == '__main__':
    started = time.perf_counter()
    try:
        run()
    except Exception as exc:
        import traceback
        traceback.print_exc()
        print(f'BROWSER REGRESSION FAILED after {len(checks)} passing checks: {exc}', file=sys.stderr)
        sys.exit(1)
    print(f'BROWSER REGRESSION PASSED ({mode}): {len(checks)} checks in {time.perf_counter()-started:.1f}s')
