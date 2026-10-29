#!/usr/bin/env python3
"""In-memory Chromium smoke; checks chart UI and OOXML round-trip, without localhost."""
import base64,json,pathlib,re,os
from playwright.sync_api import sync_playwright
ROOT=pathlib.Path(__file__).resolve().parent.parent
DIST=ROOT/'dist';TEST=pathlib.Path('/mnt/data/GridForge_v10_8_Native_Charts_Example.xlsx')
def script_content():
    imports={}
    for module in DIST.rglob('*.js'):
        source=module.read_text(encoding='utf-8')
        def repl(m):
            target=(module.parent/m.group(2)).resolve().relative_to(DIST.resolve()).as_posix()
            return m.group(1)+'gridforge/'+target+m.group(3)
        source=re.sub(r"(from\s+['\"])(\.[^'\"]+)(['\"])",repl,source)
        imports['gridforge/'+module.relative_to(DIST).as_posix()]='data:text/javascript;base64,'+base64.b64encode(source.encode()).decode()
    html=(DIST/'index.html').read_text(encoding='utf-8').replace('<link rel="stylesheet" href="./styles.css"/>','<style>'+(DIST/'styles.css').read_text(encoding='utf-8')+'</style>')
    html=html.replace('<script type="module" src="./main.js"></script>','')
    html=html.replace('</body>',f'<script type="importmap">{json.dumps({"imports":imports})}</script><script type="module">import "gridforge/main.js";</script></body>')
    return html
with sync_playwright() as p:
    kw={'headless':True,'args':['--no-sandbox']}
    if os.environ.get('GRIDFORGE_CHROMIUM'):kw['executable_path']=os.environ['GRIDFORGE_CHROMIUM']
    elif pathlib.Path('/usr/bin/chromium').exists():kw['executable_path']='/usr/bin/chromium'
    browser=p.chromium.launch(**kw)
    try:
        page=browser.new_page(viewport={'width':1440,'height':900},device_scale_factor=1)
        errors=[];page.on('pageerror',lambda err:errors.append(str(err)))
        page.set_content(script_content(),wait_until='load',timeout=30000)
        page.locator('[data-ref="D4"]').wait_for(timeout=15000)
        page.locator('[data-ribbon="Insert"]').click()
        page.locator('[data-chart-kind="stacked-column"]').first.click()
        page.locator('#chart-dialog').wait_for(state='visible')
        page.locator('#chart-title').fill('Quarterly revenue')
        page.locator('#chart-range').fill('A3:D11')
        page.locator('#chart-legend').select_option('right')
        page.locator('#chart-xaxis').fill('Product')
        page.locator('#chart-yaxis').fill('Sales amount')
        page.locator('#chart-labels').check()
        assert not page.locator('#chart-submit').is_disabled(),'chart editor incorrectly disabled'
        page.locator('#chart-submit').click()
        page.locator('.worksheet-chart').first.wait_for(timeout=10000)
        svg=page.locator('.worksheet-chart-svg svg').first
        assert svg.count()==1
        assert 'Quarterly revenue' in page.locator('.worksheet-chart-head').first.inner_text()
        page.screenshot(path=str(ROOT/'docs/v10_8_excel_charts_ui.png'),full_page=False)
        print('PASS: Insert -> stacked chart, title, legend, axis labels, value labels, live worksheet chart')
        contents=base64.b64encode(TEST.read_bytes()).decode()
        res=page.evaluate('''async b64=>{
          const binary=atob(b64),bytes=new Uint8Array(binary.length);for(let i=0;i<binary.length;i++)bytes[i]=binary.charCodeAt(i);
          const {importXlsx,exportXlsx}=await import('gridforge/engine/xlsx.js');
          const {document,warnings}=await importXlsx(bytes);
          const {zipRead}=await import('gridforge/engine/zip.js');const output=await zipRead(exportXlsx(document));
          const chartParts=[...output.keys()].filter(n=>n.startsWith('xl/charts/chart') && n.endsWith('.xml'));
          return {charts:document.active.tools.charts.map(c=>({kind:c.kind,range:c.range,title:c.title,legend:c.legend,xAxisTitle:c.xAxisTitle,yAxisTitle:c.yAxisTitle})),roundTripParts:chartParts.length,warnings};
        }''',contents)
        assert len(res['charts'])==9 and res['roundTripParts']==9,res
        assert all(c['xAxisTitle']=='Month' and c['yAxisTitle']=='Amount' for c in res['charts'] if c['kind'] not in ('pie','doughnut')),res
        assert {x['kind'] for x in res['charts']}=={'column','bar','line','area','pie','doughnut','scatter','stacked-column','stacked-bar'},res
        print('PASS: Imported 9 native DrawingML chart objects, including stacked, doughnut and scatter')
        print('Warnings:',res['warnings'][:2])
        assert not errors,errors
        print('PASS: Chromium no JS page errors')
    finally:browser.close()
