#!/usr/bin/env python3
"""Fast isolated Chromium checks for v10.4 tabs. Never touches user browser storage."""
import base64,json,os,re,sys,pathlib
from playwright.sync_api import sync_playwright
ROOT=pathlib.Path(__file__).resolve().parent.parent
p=ROOT/'dist'
with sync_playwright() as x:
 options={'headless':True,'args':['--no-sandbox']}
 executable=os.environ.get('GRIDFORGE_CHROMIUM')
 if executable: options['executable_path']=executable
 try: browser=x.chromium.launch(**options)
 except Exception:
  if os.name=='nt' or not pathlib.Path('/usr/bin/chromium').exists():raise
  browser=x.chromium.launch(headless=True,executable_path='/usr/bin/chromium',args=['--no-sandbox'])
 try:
  page=browser.new_page(viewport={'width':1440,'height':900})
  errors=[];page.on('pageerror',lambda e:errors.append(str(e)))
  sources={}
  for module in p.rglob('*.js'):
   source=module.read_text(encoding='utf-8')
   def repl(m):return m.group(1)+'gridforge/'+(module.parent/m.group(2)).resolve().relative_to(p.resolve()).as_posix()+m.group(3)
   source=re.sub(r"(from\s+['\"])(\.[^'\"]+)(['\"])",repl,source)
   sources['gridforge/'+module.relative_to(p).as_posix()]='data:text/javascript;base64,'+base64.b64encode(source.encode()).decode()
  html=(p/'index.html').read_text(encoding='utf-8')
  html=html.replace('<link rel="stylesheet" href="./styles.css"/>','<style>'+(p/'styles.css').read_text(encoding='utf-8')+'</style>')
  html=html.replace('<script type="module" src="./main.js"></script>','')
  html=html.replace('</body>',f'<script type="importmap">{json.dumps({"imports":sources})}</script><script type="module">import "gridforge/main.js";</script></body>')
  page.set_content(html,wait_until='load',timeout=15000)
  page.locator('[data-ref="D4"]').wait_for(timeout=10000)
  checks=0
  def verify(desc,cond):
   global checks
   assert cond,desc
   checks+=1;print(f'PASS {checks:02} {desc}',flush=True)
  verify('Home tab selected',page.locator('[data-ribbon="Home"]').get_attribute('aria-selected')=='true')
  verify('Font controls visible on Home',page.locator('#fmt-bold').is_visible())
  verify('File export hidden on Home',not page.locator('#btn-json').is_visible())
  for tab,ctl in [('File','#btn-xlsx'),('Insert','#btn-autosum'),('Formulas','#btn-autosum'),('Data','#btn-sort-custom'),('Review','#btn-find'),('View','#btn-freeze-row'),('Automate','#btn-automate')]:
   page.locator(f'[data-ribbon="{tab}"]').click()
   verify(f'{tab} has the expected command',page.locator(ctl).is_visible())
  page.locator('[data-ribbon="More ▾"]').count()  # canonical All value below
  page.locator('[data-ribbon="All"]').click()
  verify('All shows advanced commands',page.locator('#btn-data-studio').is_visible() and page.locator('#btn-conditional').is_visible())
  page.locator('[data-ribbon="Home"]').click()
  page.locator('#btn-collapse-ribbon').click()
  verify('Ribbon collapses, expanding grid',page.evaluate('document.documentElement.dataset.ribbonCollapsed')=='true' and not page.locator('#btn-undo').is_visible())
  page.keyboard.press('Control+F1')
  verify('Ctrl+F1 expands ribbon',page.evaluate('document.documentElement.dataset.ribbonCollapsed')=='false' and page.locator('#btn-undo').is_visible())
  page.locator('[data-ribbon="Home"]').focus();page.keyboard.press('ArrowRight')
  verify('Arrow keys navigate tab list',page.locator('[data-ribbon="Insert"]').get_attribute('aria-selected')=='true')
  page.locator('[data-ribbon="Formulas"]').click()
  page.locator('#function-category').select_option('External & Connected')
  verify('Ten connected names visible with honest warnings',page.locator('[data-function]').count()==10 and 'requires' in page.locator('#sidebar-body').inner_text().lower())
  page.locator('#name-box').fill('Z150');page.locator('#name-box').press('Enter')
  page.locator('#formula').fill('=WEBSERVICE("https://example.org/")');page.locator('#formula').press('Enter')
  verify('Connected function gives a clear offline error, not a fake result',page.locator('[data-ref="Z150"]').inner_text()=='#CONNECT!')
  page.locator('#name-box').fill('D4');page.locator('#name-box').press('Enter')
  page.locator('[data-ribbon="Home"]').click()
  page.locator('#theme-switch').click()
  verify('Theme switch still works',page.evaluate('document.documentElement.dataset.theme')=='dark')
  page.locator('#theme-switch').click()
  page.locator('#workspace-switch').select_option('guided')
  verify('Guided workspace preserved',page.locator('#guided-home').is_visible())
  page.locator('#workspace-switch').select_option('excel')
  page.locator('[data-ribbon="Home"]').click()
  if page.locator('#sidebar').is_visible():page.locator('#btn-insights').click()
  page.screenshot(path=str(ROOT/'docs'/'preview_v10_4_excel_ribbon.png'))
  verify('No JavaScript page errors',not errors)
  print(f'V10.4 RIBBON SMOKE PASSED {checks} checks, screenshot saved.',flush=True)
 finally:browser.close()
