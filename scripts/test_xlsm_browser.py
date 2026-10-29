#!/usr/bin/env python3
"""Browser-level XLSM preservation tests; isolated in an offline browser profile.
No Microsoft Excel is required for package-preservation tests.
"""
from pathlib import Path
from playwright.sync_api import sync_playwright
import base64
import json
import re

ROOT = Path(__file__).resolve().parent.parent
DIST = ROOT / 'dist'


def main():
    modules = {}
    for module in DIST.rglob('*.js'):
        relative = module.relative_to(DIST).as_posix()
        source = module.read_text(encoding='utf-8')
        def rewrite(match):
            target = (module.parent / match.group(2)).resolve().relative_to(DIST.resolve()).as_posix()
            return match.group(1) + 'gridforge/' + target + match.group(3)
        source = re.sub(r"(from\s+['\"])(\.[^'\"]+)(['\"])", rewrite, source)
        modules['gridforge/' + relative] = 'data:text/javascript;base64,' + base64.b64encode(source.encode()).decode()
    html = '<!doctype html><html><head><meta charset="utf-8"></head><body><script type="importmap">' + json.dumps({'imports': modules}) + '</script></body></html>'
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, executable_path='/usr/bin/chromium' if Path('/usr/bin/chromium').exists() else None, args=['--no-sandbox'])
        try:
            page = browser.new_page()
            page.set_content(html)
            result = page.evaluate('''async () => {
              const {SpreadsheetDocument} = await import('gridforge/engine/document.js');
              const {exportXlsx} = await import('gridforge/engine/xlsx.js');
              const {zipReadBytes,zipWriteBytes} = await import('gridforge/engine/zip.js');
              const {exportEditedXlsm} = await import('gridforge/engine/xlsm.js');
              const enc=new TextEncoder(),dec=new TextDecoder();
              const doc=new SpreadsheetDocument();
              doc.renameSheet(doc.active.id,'Sales');
              doc.active.book.applyEdits([{ref:'A1',raw:'Item'},{ref:'B1',raw:'Amount'},{ref:'A2',raw:'Widget'},{ref:'B2',raw:'15'},{ref:'C2',raw:'=B2*2'}]);
              const zip=await zipReadBytes(exportXlsx(doc));
              let types=dec.decode(zip.get('[Content_Types].xml'));
              types=types.replace('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml','application/vnd.ms-excel.sheet.macroEnabled.main+xml');
              types=types.replace('</Types>','<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>');
              zip.set('[Content_Types].xml',enc.encode(types));
              let rels=dec.decode(zip.get('xl/_rels/workbook.xml.rels'));
              rels=rels.replace('</Relationships>','<Relationship Id="rId99" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>');
              zip.set('xl/_rels/workbook.xml.rels',enc.encode(rels));
              const vba=new Uint8Array([0xd0,0xcf,0x11,0xe0,0xa1,0xb1,0x1a,0xe1,35,24]);
              const illustration=new Uint8Array([137,80,78,71,13,10,26,10,42]);
              zip.set('xl/vbaProject.bin',vba);zip.set('xl/media/test.png',illustration);
              const source=zipWriteBytes(Object.fromEntries(zip));
              const baseline=SpreadsheetDocument.fromJSON(doc.toJSON());
              doc.active.book.applyEdits([{ref:'B2',raw:'120'},{ref:'D2',raw:'=Amount+10'},{ref:'A2',raw:'Edited'}]);
              doc.active.book.setDefinedNames([{name:'Amount',ref:'B2',comment:'Macro-preserving name test'}]);
              const result=await exportEditedXlsm(source,baseline,doc);
              const output=await zipReadBytes(result.bytes);
              const worksheet=dec.decode(output.get('xl/worksheets/sheet1.xml'));
              const assertion=(ok,message)=>{if(!ok)throw Error(message);};
              assertion(result.editedCells===3,'Expected exactly 3 cell edits');assertion(result.editedNames===1,'Expected one defined name update');
              const listInvalid=[];for(const [name,b] of output) if(name.endsWith('.xml')||name.endsWith('.rels')){const xm=new DOMParser().parseFromString(dec.decode(b),'application/xml');if(xm.getElementsByTagName('parsererror').length)listInvalid.push({name,start:dec.decode(b).slice(0,180),error:xm.getElementsByTagName('parsererror')[0].textContent.slice(0,200)});}
              if(listInvalid.length)throw Error(JSON.stringify(listInvalid));
              const reimport=(await (await import('gridforge/engine/xlsx.js')).importXlsx(result.bytes)).document;
              assertion(reimport.active.book.getRaw('B2')==='120','Edited numeric cell failed XLSM re-import');
              assertion(reimport.active.book.getRaw('A2')==='Edited','Edited text cell failed XLSM re-import');
              assertion(reimport.active.book.getRaw('D2')==='=Amount+10','Edited formula failed XLSM re-import');
              assertion(reimport.active.book.getValue('D2').value===130,'Edited named formula failed XLSM calculation');
              assertion(reimport.active.book.getDefinedNames().some(n=>n.name==='Amount'&&n.ref==='B2'),'Defined name failed XLSM round-trip');
              assertion(worksheet.includes('<v>120</v>'),'Changed number missing');
              assertion(worksheet.includes('Amount+10'),'Changed formula missing');
              assertion(worksheet.includes('Edited'),'Changed string missing');
              assertion([...vba].join(',')===[...output.get('xl/vbaProject.bin')].join(','),'VBA binary was changed');
              assertion([...illustration].join(',')===[...output.get('xl/media/test.png')].join(','),'Media binary was changed');
              assertion(dec.decode(zip.get('xl/_rels/workbook.xml.rels'))===dec.decode(output.get('xl/_rels/workbook.xml.rels')),'VBA relationship was changed');
              assertion(dec.decode(output.get('xl/workbook.xml')).includes('fullCalcOnLoad="1"'),'Calculation was not marked dirty');
              let rejection='';
              const changedSheet=SpreadsheetDocument.fromJSON(doc.toJSON());changedSheet.renameSheet(changedSheet.active.id,'Wrong');
              try{await exportEditedXlsm(source,baseline,changedSheet);}catch(e){rejection=String(e.message);}
              assertion(rejection.includes('Adding, deleting'),'Sheet-structure changes were not rejected');
              const noChange=await exportEditedXlsm(source,baseline,baseline);
              const originalParts=await zipReadBytes(noChange.bytes);
              assertion([...originalParts.get('xl/vbaProject.bin')].join(',')===[...vba].join(','),'No-change preserved VBA mismatch');
              return {changedCells:result.editedCells,changedNames:result.editedNames,preservedParts:result.preservedParts,outputSize:result.bytes.length,warningCount:result.warnings.length};
            }''')
            print('PASS XLSM edited values, formulas, new cells, original VBA, binary parts, calc flags, structure safeguards:', json.dumps(result))
        finally:
            browser.close()


if __name__ == '__main__':
    main()
