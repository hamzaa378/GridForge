#!/usr/bin/env python3
"""v10.9 offline Chromium UI, table/chart, and native Excel XLSX round-trip smoke."""
import base64,json,pathlib,re,os
from playwright.sync_api import sync_playwright
ROOT=pathlib.Path(__file__).resolve().parent.parent
DIST=ROOT/'dist'
def document_html():
    imports={}
    for module in DIST.rglob('*.js'):
        source=module.read_text(encoding='utf-8')
        def rewrite(match):
            target=(module.parent/match.group(2)).resolve().relative_to(DIST.resolve()).as_posix()
            return match.group(1)+'gridforge/'+target+match.group(3)
        source=re.sub(r"(from\s+['\"])(\.[^'\"]+)(['\"])",rewrite,source)
        imports['gridforge/'+module.relative_to(DIST).as_posix()]='data:text/javascript;base64,'+base64.b64encode(source.encode()).decode()
    html=(DIST/'index.html').read_text(encoding='utf-8')
    html=html.replace('<link rel="stylesheet" href="./styles.css"/>','<style>'+(DIST/'styles.css').read_text(encoding='utf-8')+'</style>')
    html=html.replace('<script type="module" src="./main.js"></script>','')
    return html.replace('</body>',f'<script type="importmap">{json.dumps({"imports":imports})}</script><script type="module">import "gridforge/main.js";</script></body>')
with sync_playwright() as p:
    options={'headless':True,'args':['--no-sandbox']}
    if os.environ.get('GRIDFORGE_CHROMIUM'):options['executable_path']=os.environ['GRIDFORGE_CHROMIUM']
    elif pathlib.Path('/usr/bin/chromium').exists():options['executable_path']='/usr/bin/chromium'
    browser=p.chromium.launch(**options)
    try:
        page=browser.new_page(viewport={'width':1536,'height':900},device_scale_factor=1)
        errors=[];page.on('pageerror',lambda error:errors.append(str(error)))
        page.set_content(document_html(),wait_until='load',timeout=30000)
        page.locator('[data-ref="A3"]').wait_for(timeout=15000)
        page.locator('#btn-format-table').click()
        page.locator('#table-dialog').wait_for(state='visible')
        page.locator('#table-range').fill('A3:G11')
        page.locator('#table-name').fill('QuarterlySales')
        page.locator('[data-table-style="TableStyleMedium4"]').click()
        page.locator('#table-first').check()
        page.screenshot(path=str(ROOT/'docs/v10_9_table_style_gallery.png'),full_page=False)
        page.locator('#table-apply').click()
        assert page.locator('#table-dialog').is_hidden()
        color=page.locator('[data-ref="A3"]').get_attribute('style')
        assert 'background-color' in color,color
        print('PASS: Home -> Format as Table, style gallery, named table, banding, grid styling')
        page.locator('[data-ribbon="Table Design"]').click()
        page.locator('#btn-table-to-chart').click()
        page.locator('#chart-dialog').wait_for(state='visible')
        assert page.locator('#chart-range').input_value()=='A3:G11'
        page.locator('#chart-kind').select_option('combo')
        page.locator('#chart-theme').select_option('forest')
        page.locator('#chart-title').fill('Sales and Profit')
        assert page.locator('#chart-submit').is_enabled()
        page.locator('#chart-submit').click()
        assert page.locator('.worksheet-chart').count()==1
        page.screenshot(path=str(ROOT/'docs/v10_9_excel_table_chart.png'),full_page=False)
        print('PASS: Table Design -> Chart from Table -> Combo + Forest theme, worksheet SVG chart')
        page.locator('[data-ribbon="File"]').click()
        with page.expect_download() as emitted:
            page.locator('#btn-json').click()
        json_doc=json.loads(pathlib.Path(emitted.value.path()).read_text(encoding='utf-8'))
        assert len(json_doc['sheets'][0]['tools']['tables'])==1,json_doc
        assert json_doc['sheets'][0]['tools']['tables'][0]['name']=='QuarterlySales'
        with page.expect_download() as emitted:
            page.locator('#btn-xlsx').click()
        xlsx_bytes=pathlib.Path(emitted.value.path()).read_bytes()
        result=page.evaluate('''async encoded=>{
          const {importXlsx}=await import('gridforge/engine/xlsx.js');
          const {zipRead}=await import('gridforge/engine/zip.js');
          const binary=atob(encoded),bytes=Uint8Array.from(binary,ch=>ch.charCodeAt(0));
          const parts=await zipRead(bytes),reimport=await importXlsx(bytes);
          return {exportHasTable:parts.has('xl/tables/table1.xml'),importedTables:reimport.document.active.tools.tables,importedCharts:reimport.document.active.tools.charts,warnings:reimport.warnings};
        }''',base64.b64encode(xlsx_bytes).decode())
        assert result['exportHasTable'],result
        assert len(result['importedTables'])==1,result
        assert result['importedTables'][0]['name']=='QuarterlySales',result
        assert len(result['importedCharts'])==1,result
        assert result['importedCharts'][0]['kind']=='combo',result
        assert result['importedCharts'][0].get('theme')=='forest',result
        print('PASS: GridForge JSON -> native XLSX table + combo chart -> GridForge import round-trip')

        demo=base64.b64encode((ROOT/'examples/Excel_Migration_Demo_v10_9.xlsx').read_bytes()).decode()
        full=page.evaluate('''async encoded=>{
          const {importXlsx}=await import('gridforge/engine/xlsx.js');
          const bytes=Uint8Array.from(atob(encoded),ch=>ch.charCodeAt(0));
          const {document,warnings}=await importXlsx(bytes);
          return {sheets:document.sheets.length,charts:document.sheets.reduce((n,s)=>n+s.tools.charts.length,0),tables:document.sheets.reduce((n,s)=>n+s.tools.tables.length,0),kinds:document.sheets[0].tools.charts.map(x=>x.kind),warnings};
        }''',demo)
        assert full['sheets']==2 and full['charts']==13 and full['tables']==11,full
        assert all(kind in full['kinds'] for kind in ['combo','radar','percent-bar','percent-column']),full
        print('PASS: 13-chart, 11-table XLSX demo imports both sheets with all model objects')

        page.locator('[data-ribbon="Home"]').click()
        page.locator('[data-ref="B5"]').click()
        page.locator('#fmt-font').select_option('Cambria')
        page.locator('#fmt-size').select_option('28')
        page.locator('[data-color-swatch="#fff2ba"]').click()
        print('PASS: Excel-like font selection, size and one-click fill color')
        assert not errors,errors
        print('PASS: Chromium no page errors')
    finally:browser.close()
