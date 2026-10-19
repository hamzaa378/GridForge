from playwright.sync_api import sync_playwright
from pathlib import Path
import os
import base64,re,json
root=Path(__file__).resolve().parent.parent/'dist'
imports={}
for file in root.rglob('*.js'):
 rel=file.relative_to(root).as_posix();source=file.read_text()
 def rewrite(match):
  target=(file.parent/match.group(2)).resolve().relative_to(root.resolve()).as_posix()
  return match.group(1)+'gridforge/'+target+match.group(3)
 source=re.sub(r'''(from\s+['"])(\.[^'"]+)(['"])''',rewrite,source)
 imports['gridforge/'+rel]='data:text/javascript;base64,'+base64.b64encode(source.encode()).decode()
html=(root/'index.html').read_text().replace('<link rel="stylesheet" href="./styles.css"/>','<style>'+(root/'styles.css').read_text()+'</style>').replace('<script type="module" src="./main.js"></script>','')
html=html.replace('</body>',f'<script type="importmap">{json.dumps({"imports":imports})}</script><script type="module">import "gridforge/main.js";</script></body>')
with sync_playwright() as p:
 browser=p.chromium.launch(headless=True,args=['--no-sandbox'],**({'executable_path':os.environ['GRIDFORGE_CHROMIUM']} if os.environ.get('GRIDFORGE_CHROMIUM') else {}))
 page=browser.new_page(viewport={'width':1440,'height':900})
 errors=[];page.on('pageerror',lambda e: errors.append(str(e)))
 page.set_content(html,wait_until='load',timeout=12000);page.wait_for_timeout(150)
 def cell(ref):return page.locator(f'[data-ref="{ref}"]')
 assert cell('D4').inner_text()=='4000'
 page.locator('#name-box').fill('B4');page.locator('#name-box').press('Enter')
 page.locator('#formula').fill('225');page.locator('#formula').press('Enter')
 assert cell('D4').inner_text()=='7200',cell('D4').inner_text()
 page.locator('#btn-undo').click();assert cell('D4').inner_text()=='4000'
 # format -> undo -> redo
 cell('D4').click();page.locator('#fmt-bold').click();assert cell('D4').evaluate('(e)=>e.style.fontWeight')=='800'
 page.locator('#btn-undo').click();assert cell('D4').evaluate('(e)=>e.style.fontWeight')!='800'
 page.locator('#btn-redo').click();assert cell('D4').evaluate('(e)=>e.style.fontWeight')=='800'
 # sheet add / independently edit / return
 page.locator('#btn-add-sheet').click();assert page.locator('[data-sheet-id]').count()==2
 page.locator('#formula').fill('42');page.locator('#formula').press('Enter')
 assert cell('A1').inner_text()=='42'
 page.locator('[data-sheet-id="s1"]').click();assert cell('D4').inner_text()=='4000'
 page.locator('[data-sheet-id="s2"]').click();assert cell('A1').inner_text()=='42'
 page.locator('[data-sheet-id="s1"]').click()
 # find replace, with formula protection
 page.locator('#btn-find').click();page.locator('#find-query').fill('P-103');page.locator('#btn-find-all').click()
 assert page.locator('.search-result').count()==2
 page.locator('#close-find').click()
 # clipboard using native ClipboardEvent and custom MIME
 cell('D4').click();page.locator('#grid-scroll').focus()
 page.evaluate('''() => {window.__gCopy = new DataTransfer(); document.dispatchEvent(new ClipboardEvent('copy',{clipboardData:window.__gCopy,bubbles:true,cancelable:true}));}''')
 assert page.evaluate('window.__gCopy.getData("application/x-gridforge-cells").length')>0
 cell('H4').click();page.locator('#grid-scroll').focus()
 page.evaluate('''() => document.dispatchEvent(new ClipboardEvent('paste',{clipboardData:window.__gCopy,bubbles:true,cancelable:true}))''')
 page.locator('#name-box').fill('H4');page.locator('#name-box').press('Enter')
 raw=page.locator('#formula').input_value()
 assert raw=='=F4*G4',raw
 page.locator('#btn-undo').click();assert page.locator('#formula').input_value()==''
 # analyze range by click and shift-click
 cell('A4').click();cell('B11').click(modifiers=['Shift']);page.locator('[data-mode="analyze"]').click()
 assert page.locator('.quick-chart').count()==1
 print('analysis:',page.locator('.analysis-stats').inner_text().replace('\n',' '))
 # Optional screenshot: uncomment to capture a demo image
 # page.screenshot(path=str(root.parent/'docs'/'preview_analytics_v2.png'),full_page=True)
 # sort selected rectangle then undo
 page.locator('#btn-sort-desc').click()
 assert cell('A4').inner_text()=='P-107',cell('A4').inner_text()
 page.locator('#btn-undo').click();assert cell('A4').inner_text()=='P-100'
 # JSON export is a v2 multi-sheet snapshot; real localStorage is unavailable on about:blank.
 with page.expect_download(timeout=5000) as info:page.locator('#btn-json').click()
 export=info.value
 payload=json.loads(export.path().read_text())
 assert len(payload['sheets'])==2 and payload['version']==2
 assert payload['sheets'][1]['workbook']['cells']['A1']=='42'
 assert payload['sheets'][0]['workbook']['styles']['D4']['bold'] is True
 print('RESULT: UI checks pass','visible cells:',page.locator('.cell').count(),'sheets:',len(payload['sheets']),'errors:',errors)
 assert not errors,errors
 browser.close()
