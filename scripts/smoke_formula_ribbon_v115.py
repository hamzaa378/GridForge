#!/usr/bin/env python3
"""Focused browser checks for the v11.5 Formula ribbon and auditing workflows."""
from __future__ import annotations
import json, os, socket, subprocess, time
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
        cells={'A1':'10','A2':'20','B1':'=SUM(A1:A2)','C1':'=B1*2','D1':'10','D2':'20'}
        doc={'format':'gridforge','version':3,'activeId':'s1','sheets':[{'id':'s1','name':'Sheet1','workbook':{'version':2,'rows':100,'cols':26,'cells':cells,'styles':{}},'tools':{'freezeRows':0,'freezeCols':0,'filter':None,'conditional':[],'charts':[],'tables':[],'columnWidths':{},'rowHeights':{}}}]}
        with sync_playwright() as p:
            browser=p.chromium.launch(headless=True,args=['--no-sandbox']);context=browser.new_context(viewport={'width':1500,'height':950})
            context.add_init_script("localStorage.setItem('gridforge:document:v2',"+json.dumps(json.dumps(doc))+');')
            page=context.new_page();page.set_default_timeout(8000);errors=[];page.on('pageerror',lambda e:errors.append(str(e)));page.goto(url,wait_until='load')
            page.locator('[data-ribbon="Formulas"]').click()
            assert page.locator('#formula-ribbon').is_visible()
            preview_dir=ROOT/'work';preview_dir.mkdir(exist_ok=True)
            page.screenshot(path=str(preview_dir/'GridForge_v11_5_Formula_Ribbon_Preview.png'))
            page.locator('#formula-ribbon [aria-label="Function categories"] .formula-menu summary').first.click();page.locator('[data-function="IF"]').click()
            assert page.locator('#formula').input_value()=='=IF()'
            print('PASS Formula ribbon category menu inserts functions into formula bar')
            page.locator('#btn-insert-function').click();page.locator('#function-dialog-search').fill('AVERAGE')
            page.locator('[data-select-function="AVERAGE"]').click();page.locator('#function-dialog-insert').click()
            assert page.locator('#formula').input_value()=='=AVERAGE()'
            print('PASS searchable Insert Function dialog')
            page.locator('[data-ref="D1"]').click();page.locator('[data-ref="D2"]').click(modifiers=['Shift'])
            page.locator('#formula-aggregate').select_option('AVERAGE');page.locator('#btn-formula-aggregate').click()
            assert page.locator('[data-ref="D3"]').inner_text()=='15'
            page.wait_for_timeout(350)
            assert page.evaluate("JSON.parse(localStorage.getItem('gridforge:document:v2')).sheets[0].workbook.cells.D3")=='=AVERAGE(D1:D2)'
            print('PASS AutoSum aggregate selector inserts a calculated AVERAGE')
            page.locator('[data-ref="C1"]').click();page.locator('#btn-trace-precedents').click()
            assert 'traced-precedent' in page.locator('[data-ref="B1"]').get_attribute('class')
            page.locator('#btn-remove-formula-arrows').click();assert 'traced-precedent' not in page.locator('[data-ref="B1"]').get_attribute('class')
            page.locator('[data-ref="A1"]').click();page.locator('#btn-trace-dependents').click()
            assert 'traced-dependent' in page.locator('[data-ref="B1"]').get_attribute('class')
            print('PASS precedent/dependent trace and clear')
            page.locator('#btn-show-formulas').click()
            assert page.locator('[data-ref="B1"]').inner_text()=='=SUM(A1:A2)'
            page.locator('#btn-show-formulas').click();assert page.locator('[data-ref="B1"]').inner_text()=='30'
            print('PASS worksheet Show Formulas toggle')
            page.locator('[data-ref="C1"]').click();page.locator('#btn-evaluate-formula').click()
            assert page.locator('#evaluate-formula-dialog').is_visible()
            assert '=B1*2' in page.locator('#evaluate-formula-content').inner_text() and '60' in page.locator('#evaluate-formula-content').inner_text()
            page.locator('[data-close-evaluate]').click()
            assert not errors,errors
            print('PASS formula evaluation details; no browser errors')
            browser.close()
    finally:
        proc.terminate()
        try:proc.wait(timeout=5)
        except subprocess.TimeoutExpired:proc.kill()
if __name__=='__main__':main()
