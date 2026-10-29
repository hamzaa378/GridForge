#!/usr/bin/env python3
"""Focused Chromium interaction checks for v11.2 Excel-style ribbon, chart, table and grid workflows."""
from __future__ import annotations
import json, os, socket, subprocess, sys, tempfile, time, zipfile
from pathlib import Path
from urllib.request import urlopen
from playwright.sync_api import sync_playwright

ROOT=Path(__file__).resolve().parent.parent
def free_port():
    with socket.socket() as s:s.bind(('127.0.0.1',0));return s.getsockname()[1]
def wait(url,proc):
    for _ in range(100):
        if proc.poll() is not None:raise RuntimeError('GridForge server exited')
        try:
            with urlopen(url,timeout=.4) as r:
                if r.status==200:return
        except Exception:time.sleep(.1)
    raise RuntimeError('GridForge server not ready')
def main():
    port=free_port();url=f'http://127.0.0.1:{port}';env={**os.environ,'PORT':str(port)}
    proc=subprocess.Popen(['node','scripts/serve.mjs'],cwd=ROOT,env=env,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
    try:
        wait(url,proc)
        cells={'A1':'Month','B1':'Revenue','C1':'Margin','D1':'Total','A2':'Jan','B2':'1000','C2':'0.1','A3':'Feb','B3':'2000','C3':'0.2'}
        table={'id':'table-metrics','name':'Metrics','range':'A1:D3','style':'TableStyleMedium2','header':True,'bandedRows':True,'bandedColumns':False,'firstColumn':False,'lastColumn':False}
        initial={'format':'gridforge','version':3,'activeId':'s1','sheets':[{'id':'s1','name':'Sheet 1','workbook':{'version':2,'rows':400,'cols':52,'cells':cells,'styles':{}},'tools':{'freezeRows':0,'freezeCols':0,'filter':None,'conditional':[],'charts':[],'tables':[table],'columnWidths':{},'rowHeights':{}}}]}
        with sync_playwright() as p:
            opts={'headless':True,'args':['--no-sandbox']}
            if os.getenv('GRIDFORGE_CHROMIUM'):opts['executable_path']=os.environ['GRIDFORGE_CHROMIUM']
            browser=p.chromium.launch(**opts);context=browser.new_context(viewport={'width':1440,'height':900},accept_downloads=True)
            context.add_init_script("localStorage.setItem('gridforge:document:v2',"+json.dumps(json.dumps(initial))+ ");")
            page=context.new_page();page.set_default_timeout(10000);errors=[];page.on('pageerror',lambda e:errors.append(str(e)));page.goto(url,wait_until='load')
            def enter(ref,value):
                page.locator(f'[data-ref="{ref}"]').click();page.locator('#formula').fill(value);page.locator('#formula').press('Enter')
            enter('D2','=[@Revenue]*[@Margin]')
            assert page.locator('[data-ref="D2"]').inner_text()=='100'
            assert page.locator('[data-ref="D3"]').inner_text()=='400'
            assert page.locator('[data-ribbon="Table Design"]').is_visible()
            print('PASS calculated-column propagation through UI')
            # Excel-style whole-column selection supports a shared resize.
            page.locator('.col-heading[data-column="1"]').click();page.locator('.col-heading[data-column="2"]').click(modifiers=['Shift'])
            heading=page.locator('.col-heading[data-column="2"]');box=heading.bounding_box();page.mouse.move(box['x']+box['width']-1,box['y']+box['height']/2);page.mouse.down();page.mouse.move(box['x']+box['width']+18,box['y']+box['height']/2,steps=4);page.mouse.up()
            widths=page.locator('.col-heading').evaluate_all('(nodes)=>[nodes.find(n=>n.dataset.column==="1").style.width,nodes.find(n=>n.dataset.column==="2").style.width]');assert widths[0]==widths[1],widths
            print('PASS multi-column selection and shared resize')
            # Drag the B column boundary; width is saved in the native document state.
            heading=page.locator('.col-heading[data-column="1"]');box=heading.bounding_box();page.mouse.move(box['x']+box['width']-1,box['y']+box['height']/2);page.mouse.down();page.mouse.move(box['x']+box['width']+35,box['y']+box['height']/2,steps=5);page.mouse.up()
            # AutoFit the same column by double-clicking its current boundary.
            heading=page.locator('.col-heading[data-column="1"]');box=heading.bounding_box();page.mouse.dblclick(box['x']+box['width']-1,box['y']+box['height']/2)
            rowhead=page.locator('.row-heading[data-row="1"]');box=rowhead.bounding_box();page.mouse.move(box['x']+box['width']/2,box['y']+box['height']-1);page.mouse.down();page.mouse.move(box['x']+box['width']/2,box['y']+box['height']+16,steps=5);page.mouse.up()
            rowhead=page.locator('.row-heading[data-row="1"]');box=rowhead.bounding_box();page.mouse.dblclick(box['x']+box['width']/2,box['y']+box['height']-1)
            print('PASS row/column drag resize and AutoFit interaction')
            # Apply a live data-bar rule to the selected numeric values.
            page.locator('[data-ref="B2"]').click();page.locator('[data-ref="B3"]').click(modifiers=['Shift']);page.locator('[data-ribbon="Data"]').click()
            page.locator('#btn-conditional').click();page.locator('#rule-op').select_option('dataBar');page.locator('#rule-color').fill('#2878c8');page.locator('#apply-conditional').click()
            assert 'linear-gradient' in page.locator('[data-ref="B3"]').evaluate('(e)=>e.style.background')
            print('PASS live data-bar rule through UI')
            # Create a combo chart and put margin on its own axis with a linear trendline.
            page.locator('[data-ref="A1"]').click();page.locator('[data-ref="C3"]').click(modifiers=['Shift']);page.locator('[data-ribbon="Insert"]').click();page.locator('[data-chart-kind="combo"]').click()
            page.locator('#chart-range').fill('A1:C3');page.locator('#chart-range').dispatch_event('input');page.locator('#chart-legend').select_option('top')
            page.locator('#chart-yaxis-secondary').fill('Margin (%)')
            page.locator('.chart-series-config').nth(1).locator('[data-series-axis]').select_option('secondary')
            page.locator('.chart-series-config').nth(1).locator('[data-trendline]').select_option('linear')
            page.locator('.chart-series-config').nth(1).locator('[data-r2]').check()
            page.locator('#chart-submit').click();page.locator('.worksheet-chart-svg svg').wait_for()
            assert 'Margin (%)' in page.locator('.worksheet-chart-svg svg').inner_html()
            assert page.locator('.worksheet-chart-svg svg rect[y="56"]').count()==2
            page.locator('#chart-toggle-error-bars').click();page.locator('.worksheet-chart-svg svg g[stroke]').first.wait_for()
            print('PASS secondary-axis chart and trendline editor')
            assert page.locator('[data-ribbon="Chart Design"]').is_visible() and page.locator('[data-ribbon="Format"]').is_visible()
            assert page.locator('#chart-context-kind').input_value()=='combo'
            page.locator('#chart-switch-data').click();page.locator('[data-ribbon="Format"]').click()
            page.locator('#chart-format-title').fill('Revenue and margin')
            assert page.locator('.worksheet-chart-head strong').all_inner_texts()[-1]=='Revenue and margin'
            page.locator('#chart-format-x-title').fill('Month')
            page.locator('#chart-format-x-title').dispatch_event('change')
            page.wait_for_function("Array.from(document.querySelectorAll('.worksheet-chart-svg svg')).some(svg=>svg.innerHTML.includes('Month'))")
            page.locator('#btn-add-sheet').click()
            assert not page.locator('[data-ribbon="Chart Design"]').is_visible() and not page.locator('[data-ribbon="Format"]').is_visible()
            page.locator('.sheet-tab').first.click()
            assert not page.locator('[data-ribbon="Chart Design"]').is_visible()
            page.locator('.worksheet-chart').click()
            assert page.locator('[data-ribbon="Chart Design"]').is_visible()
            assert page.locator('#chart-legend-position').input_value()=='top'
            page.locator('#chart-legend-position').select_option('bottom');page.locator('#chart-legend-position').select_option('top')
            print('PASS contextual chart ribbon, top legend setting and Switch Row/Column')
            # Save to XLSX, re-import, and ensure features are still in the app model.
            page.locator('[data-ribbon="File"]').click()
            with page.expect_download() as dl:page.locator('#btn-xlsx').click()
            xlsx=Path(dl.value.path())
            with zipfile.ZipFile(xlsx) as z:
                assert z.testzip() is None
                chart=z.read('xl/charts/chart1.xml').decode();sheet=z.read('xl/worksheets/sheet1.xml').decode();tablexml=z.read('xl/tables/table1.xml').decode()
                assert 'axId val="404"' in chart and 'trendlineType val="linear"' in chart and 'errValType val="percentage"' in chart and 'legendPos val="t"' in chart
                assert 'type="dataBar"' in sheet and 'customWidth="1"' in sheet
                assert 'calculatedColumnFormula' in tablexml
            page.locator('#file-picker').set_input_files({'name':'roundtrip.xlsx','mimeType':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','buffer':xlsx.read_bytes()});page.wait_for_timeout(500)
            if 'Imported' not in page.locator('#toast').inner_text():raise AssertionError('XLSX import failed: '+page.locator('#toast').inner_text()+'; browser errors='+repr(errors))
            if page.locator('#import-notes-dialog').is_visible():page.locator('#confirm-import-notes').click()
            if page.locator('.worksheet-chart-svg svg').count()!=1:raise AssertionError('Chart did not reopen; toast='+page.locator('#toast').inner_text()+'; stored='+page.evaluate("localStorage.getItem('gridforge:document:v2')"))
            imported=page.evaluate("JSON.parse(localStorage.getItem('gridforge:document:v2'))")
            assert imported['sheets'][0]['tools']['charts'][0]['switchRowsColumns'] is True
            assert imported['sheets'][0]['tools']['charts'][0]['seriesOptions'][0]['errorBars']['type']=='percentage'
            assert imported['sheets'][0]['tools']['charts'][0]['legend']=='top'
            assert page.locator('[data-ref="D2"]').inner_text()=='100'
            assert 'linear-gradient' in page.locator('[data-ref="B3"]').evaluate('(e)=>e.style.background')
            print('PASS XLSX save and reopen with chart/table/format features')
            if errors:raise AssertionError('Browser errors: '+repr(errors))
            browser.close()
    finally:
        proc.terminate()
        try:proc.wait(timeout=5)
        except subprocess.TimeoutExpired:proc.kill()
if __name__=='__main__':main()
