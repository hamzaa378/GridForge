#!/usr/bin/env python3
"""Focused browser checks for v11.6 worksheet-scoped defined names."""
from __future__ import annotations
import json, os, socket, subprocess, time, zipfile
from pathlib import Path
from urllib.request import urlopen
from playwright.sync_api import sync_playwright
ROOT=Path(__file__).resolve().parent.parent
def free_port():
    with socket.socket() as s:s.bind(('127.0.0.1',0));return s.getsockname()[1]
def main():
    port=free_port();url=f'http://127.0.0.1:{port}';env={**os.environ,'PORT':str(port)}
    proc=subprocess.Popen(['node','scripts/serve.mjs'],cwd=ROOT,env=env,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
    try:
        for _ in range(100):
            if proc.poll() is not None:raise RuntimeError('GridForge server exited')
            try:
                with urlopen(url,timeout=.4) as r:
                    if r.status==200:break
            except Exception:time.sleep(.1)
        else:raise RuntimeError('GridForge server did not start')
        cells={'A1':'Product','B1':'Revenue','A2':'North','B2':'10','A3':'South','B3':'20','C1':'=SUM(B2:B3)'}
        doc={'format':'gridforge','version':3,'activeId':'s1','sheets':[{'id':'s1','name':'Sales','workbook':{'version':2,'rows':100,'cols':26,'cells':cells,'styles':{},'definedNames':[]},'tools':{'freezeRows':0,'freezeCols':0,'filter':None,'conditional':[],'charts':[],'tables':[],'columnWidths':{},'rowHeights':{}}}]}
        with sync_playwright() as p:
            browser=p.chromium.launch(headless=True,args=['--no-sandbox']);context=browser.new_context(viewport={'width':1600,'height':950},accept_downloads=True)
            context.add_init_script("localStorage.setItem('gridforge:document:v2',"+json.dumps(json.dumps(doc))+');')
            page=context.new_page();page.set_default_timeout(8000);errors=[];page.on('pageerror',lambda e:errors.append(str(e)));page.goto(url,wait_until='load')
            page.locator('[data-ribbon="Formulas"]').click()
            assert page.locator('#btn-define-name').is_visible()
            tools=page.locator('#formula-ribbon [aria-label="Defined names"] details');tools.locator('summary').click()
            assert page.locator('#btn-name-manager').is_visible()
            page.locator('#btn-define-name').click();assert page.locator('#defined-name-ref').input_value()=='A1'
            page.locator('#defined-name-input').fill('TotalSales');page.locator('#defined-name-ref').fill('B2:B3');page.locator('#save-defined-name').click()
            page.locator('[data-ref="C1"]').click();page.locator('#formula').fill('=SUM(TotalSales)');page.locator('#formula').press('Enter')
            assert page.locator('[data-ref="C1"]').inner_text()=='30'
            print('PASS define name and calculate SUM over named range')
            if not tools.evaluate('(n)=>n.open'):tools.locator('summary').click()
            page.locator('#btn-name-manager').click();assert 'TotalSales' in page.locator('#name-manager-list').inner_text()
            preview_dir=ROOT/'work';preview_dir.mkdir(exist_ok=True)
            page.screenshot(path=str(preview_dir/'GridForge_v11_6_Name_Manager_Preview.png'))
            page.locator('[data-name-edit="TotalSales"]').click();page.locator('#defined-name-ref').fill('B2:B2');page.locator('#save-defined-name').click()
            assert page.locator('[data-ref="C1"]').inner_text()=='10'
            page.locator('#btn-undo').click();assert page.locator('[data-ref="C1"]').inner_text()=='30'
            page.locator('#btn-redo').click();assert page.locator('[data-ref="C1"]').inner_text()=='10'
            print('PASS Name Manager edit recalculates formulas and supports undo/redo')
            page.locator('[data-ref="A1"]').click();page.locator('[data-ref="B3"]').click(modifiers=['Shift'])
            if not tools.evaluate('(n)=>n.open'):tools.locator('summary').click()
            page.locator('#btn-create-names').click()
            page.locator('#create-name-top-row').check();page.locator('#create-name-left-column').check();page.locator('#apply-create-names').click()
            page.wait_for_timeout(350)
            assert 'Revenue' in page.evaluate("JSON.parse(localStorage.getItem('gridforge:document:v2')).sheets[0].workbook.definedNames.map(n=>n.name).join(',')")
            if not tools.evaluate('(n)=>n.open'):tools.locator('summary').click()
            page.locator('#btn-use-name').click();page.locator('[data-use-name="Revenue"]').click()
            assert page.locator('#formula').input_value()=='=Revenue'
            print('PASS create names from headings and Use in Formula')
            page.locator('[data-ribbon="File"]').click()
            with page.expect_download() as dl:page.locator('#btn-xlsx').click()
            xlsx=Path(dl.value.path())
            with zipfile.ZipFile(xlsx) as z:
                assert z.testzip() is None
                workbook_xml=z.read('xl/workbook.xml').decode()
                assert 'definedNames' in workbook_xml and 'localSheetId="0"' in workbook_xml and 'Revenue' in workbook_xml
            page.locator('#file-picker').set_input_files({'name':'named-ranges-roundtrip.xlsx','mimeType':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','buffer':xlsx.read_bytes()});page.wait_for_timeout(500)
            if page.locator('#import-notes-dialog').is_visible():page.locator('#confirm-import-notes').click()
            assert page.locator('[data-ref="C1"]').inner_text()=='10'
            saved=page.evaluate("JSON.parse(localStorage.getItem('gridforge:document:v2')).sheets[0].workbook.definedNames")
            assert any(n['name']=='Revenue' for n in saved),saved
            assert not errors,errors
            print('PASS XLSX named-range export/import and formula calculation')
            browser.close()
    finally:
        proc.terminate()
        try:proc.wait(timeout=5)
        except subprocess.TimeoutExpired:proc.kill()
if __name__=='__main__':main()
