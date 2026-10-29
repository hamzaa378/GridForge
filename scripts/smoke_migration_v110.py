#!/usr/bin/env python3
"""Offline Chromium smoke: worksheet table filtering, structured formulas, chart growth, and XLSX imports."""
import base64,json,pathlib,re,os
from playwright.sync_api import sync_playwright
ROOT=pathlib.Path(__file__).resolve().parent.parent
DIST=ROOT/'dist'
def document_html():
    imports={}
    for module in DIST.rglob('*.js'):
        source=module.read_text(encoding='utf-8')
        def rewrite(m):
            target=(module.parent/m.group(2)).resolve().relative_to(DIST.resolve()).as_posix()
            return m.group(1)+'gridforge/'+target+m.group(3)
        source=re.sub(r"(from\s+['\"])(\.[^'\"]+)(['\"])",rewrite,source)
        imports['gridforge/'+module.relative_to(DIST).as_posix()]='data:text/javascript;base64,'+base64.b64encode(source.encode()).decode()
    html=(DIST/'index.html').read_text(encoding='utf-8')
    html=html.replace('<link rel="stylesheet" href="./styles.css"/>','<style>'+(DIST/'styles.css').read_text(encoding='utf-8')+'</style>')
    html=html.replace('<script type="module" src="./main.js"></script>','')
    return html.replace('</body>',f'<script type="importmap">{json.dumps({"imports":imports})}</script><script type="module">import "gridforge/main.js";</script></body>')

def save_formula(page, ref, formula):
    page.locator('#name-box').fill(ref)
    page.locator('#name-box').press('Enter')
    page.locator('#formula').fill(formula)
    page.locator('#formula').press('Enter')

with sync_playwright() as p:
    options={'headless':True,'args':['--no-sandbox']}
    if os.environ.get('GRIDFORGE_CHROMIUM'):options['executable_path']=os.environ['GRIDFORGE_CHROMIUM']
    elif pathlib.Path('/usr/bin/chromium').exists():options['executable_path']='/usr/bin/chromium'
    browser=p.chromium.launch(**options)
    try:
        page=browser.new_page(viewport={'width':1500,'height':920},accept_downloads=True)
        errors=[];page.on('pageerror',lambda error:errors.append(str(error)))
        page.set_content(document_html(),wait_until='load',timeout=30000)
        page.locator('[data-ref="A3"]').wait_for(timeout=18000)
        page.locator('#btn-format-table').click()
        page.locator('#table-range').fill('A3:D11')
        page.locator('#table-name').fill('SalesData')
        page.locator('#table-apply').click()
        assert page.locator('[data-table-filter]').count()>=4
        print('PASS: native Excel-style table created with header filter dropdown buttons')
        page.locator('[data-table-filter="1"]').first.click()
        page.locator('#filter-dialog').wait_for(state='visible')
        assert page.locator('#filter-values input').count()>1
        page.locator('#filter-select-none').click()
        page.locator('#filter-values input').first.check()
        page.screenshot(path=str(ROOT/'docs/v11_0_filter_dropdown.png'),full_page=False)
        page.locator('#apply-filter').click()
        print('PASS: unique-value checkbox filter from table header')
        # Clear to restore all rows, then create a source-linked chart.
        page.locator('[data-ribbon="Data"]').click()
        page.locator('#btn-clear-filter').click()
        page.locator('[data-ribbon="Home"]').click()
        page.locator('[data-ref="A4"]').click()
        page.locator('[data-ribbon="Table Design"]').click()
        page.locator('#btn-table-to-chart').click()
        assert page.locator('#chart-range').input_value()=='A3:D11'
        page.locator('#chart-submit').click()
        assert page.locator('.worksheet-chart').count()==1
        save_formula(page,'H2','=SUM(SalesData[Units])')
        before=page.locator('[data-ref="H2"]').inner_text()
        assert before and not before.startswith('#'),before
        print('PASS: =SUM(SalesData[Units]) evaluated in grid')
        save_formula(page,'A12','NewSKU')
        save_formula(page,'C12','25')
        page.locator('[data-ribbon="Table Design"]').click()
        page.locator('#btn-table-edit').click()
        assert page.locator('#table-range').input_value()=='A3:D12'
        page.locator('#close-table').click()
        page.locator('.worksheet-chart [data-chart-action="edit"]').click()
        assert page.locator('#chart-range').input_value()=='A3:D12'
        page.locator('#close-chart').click()
        print('PASS: table grows into row 12 and linked chart source grows with it')
        page.screenshot(path=str(ROOT/'docs/v11_0_live_table_chart.png'),full_page=False)
        # XLSX exported with native expanded table and chart; round-trip import in browser.
        page.locator('[data-ribbon="File"]').click()
        with page.expect_download() as received:page.locator('#btn-xlsx').click()
        contents=pathlib.Path(received.value.path()).read_bytes()
        result=page.evaluate('''async encoded=>{
          const {importXlsx}=await import('gridforge/engine/xlsx.js');
          const {zipRead}=await import('gridforge/engine/zip.js');
          const bytes=Uint8Array.from(atob(encoded),c=>c.charCodeAt(0));
          const archive=await zipRead(bytes),result=await importXlsx(bytes),sheet=result.document.active;
          return {range:sheet.tools.tables[0]?.range,chartRange:sheet.tools.charts[0]?.range,structuredValue:sheet.book.getValue('H2').value,hasTable:archive.has('xl/tables/table1.xml'),warnings:result.warnings};
        }''',base64.b64encode(contents).decode())
        assert result['range']=='A3:D12',result
        assert result['chartRange']=='A3:D12',result
        assert isinstance(result['structuredValue'],(int,float)),result
        assert result['hasTable'],result
        print('PASS: native XLSX table/chart/structured formula imported and evaluated in Chromium')
        advanced=page.evaluate("""async()=>{
          const {Workbook}=await import('gridforge/engine/workbook.js');
          const {SpreadsheetDocument}=await import('gridforge/engine/document.js');
          const {importXlsx,exportXlsx}=await import('gridforge/engine/xlsx.js');
          const {zipRead}=await import('gridforge/engine/zip.js');
          const book=new Workbook(30,12);
          book.applyEdits([{ref:'A1',raw:'Region'},{ref:'B1',raw:'Revenue'},
            {ref:'A2',raw:'North'},{ref:'B2',raw:'100'},
            {ref:'A3',raw:'South'},{ref:'B3',raw:'200'},
            {ref:'A4',raw:'North'},{ref:'B4',raw:'300'}]);
          const d=new SpreadsheetDocument(book);
          d.active.tools.tables=[{id:'table-report',name:'RegionSales',range:'A1:B4',style:'TableStyleMedium2',header:true,bandedRows:true,bandedColumns:false,firstColumn:false,lastColumn:false}];
          book.setTableContext(d.active.tools.tables);
          book.setCell('E1','=SUM(RegionSales[Revenue])');
          d.active.tools.filter={startRow:0,endRow:3,startCol:0,endCol:1,column:0,query:'',criteria:[{column:0,values:['North']},{column:1,values:['100','300']}]};
          d.active.tools.conditional=[{start:'B2',end:'B4',operator:'scale',threshold:0,startColor:'#fffafa',color:'#147755'}];
          const bytes=exportXlsx(d),parts=await zipRead(bytes),round=await importXlsx(bytes);
          const imported=round.document.active;
          const formulas=imported.book.getValue('E1');
          return {table:imported.tools.tables[0],filter:imported.tools.filter,condition:imported.tools.conditional[0],formula:formulas.value,error:formulas.error,tablePart:parts.get('xl/tables/table1.xml'),sheetPart:parts.get('xl/worksheets/sheet1.xml'),warnings:round.warnings};
        }""")
        assert advanced['table']['name']=='RegionSales',advanced
        assert advanced['formula']==600,advanced
        assert len(advanced['filter']['criteria'])==2,advanced
        assert advanced['condition']['operator']=='scale',advanced
        assert 'type="colorScale"' in advanced['sheetPart'],advanced
        assert '<filter val="North"/>' in advanced['tablePart'],advanced
        print('PASS: multi-column filters, color scales and table formulas round-trip through native Excel XLSX')
        fixture=base64.b64encode((ROOT/'examples/GridForge_v11_0_Table_Migration_Demo.xlsx').read_bytes()).decode()
        actual=page.evaluate("""async encoded=>{
          const {importXlsx}=await import('gridforge/engine/xlsx.js');
          const binary=Uint8Array.from(atob(encoded),c=>c.charCodeAt(0));
          const d=(await importXlsx(binary)).document;const sheet=d.active;
          return {total:sheet.book.getValue('F2').value,row:sheet.book.getValue('D7').value,charts:sheet.tools.charts.length,filter:sheet.tools.filter,scale:sheet.tools.conditional[0]?.operator,table:sheet.tools.tables[0]?.range};
        }""",fixture)
        assert actual['total']==3085 and actual['row']==600,actual
        assert actual['charts']==2 and actual['table']=='A1:D7',actual
        assert len(actual['filter']['criteria'])==2 and actual['scale']=='scale',actual
        print('PASS: packaged Excel demo imported with intact table, formulas, charts, filters and color scale')
        assert not errors,errors
        print('PASS: no uncaught Chromium page errors')
    finally:browser.close()
