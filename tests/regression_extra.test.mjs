/** Additional regression canaries: these defend core invariants against future UI upgrades. */
import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook} from '../dist/engine/workbook.js';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {sortRangeTransaction} from '../dist/engine/operations.js';
import {exportXlsx} from '../dist/engine/xlsx.js';
import {zipRead} from '../dist/engine/zip.js';
import {shiftFormula} from '../dist/engine/parser.js';

const create = cells => {
  const w = new Workbook(400,52);
  w.applyEdits(Object.entries(cells).map(([ref,raw])=>({ref,raw:String(raw)})),false);
  return w;
};
test('REGRESSION: deterministic edits only recalculate affected dependencies', () => {
  const w=create({A1:'1',B1:'=A1*2',C1:'=B1+10',X1:'=999*2'});
  for (let n=2;n<=150;n++) {
    const result=w.setCell('A1',String(n));
    assert.deepEqual(new Set(result.affected),new Set(['A1','B1','C1']));
    assert.equal(result.recalculated,3);
    assert.equal(w.getValue('C1').value,2*n+10);
    assert.equal(w.getValue('X1').value,1998);
  }
});
test('REGRESSION: after undo, a new edit clears obsolete redo history',()=>{
  const w=create({A1:'1'});
  w.setCell('A1','2');w.setCell('A1','3');assert.ok(w.undo());
  assert.equal(w.getRaw('A1'),'2');w.setCell('A1','7');
  assert.equal(w.canRedo,false);assert.equal(w.redo(),false);
  assert.equal(w.getRaw('A1'),'7');
});
test('REGRESSION: bad style inside transaction cannot partially change cell data',()=>{
  const w=create({A1:'5'});
  assert.throws(()=>w.applyTransaction([{ref:'A1',raw:'25'}],[{ref:'B1',style:{background:'unsafe-css'}}]),/Invalid background/);
  assert.equal(w.getRaw('A1'),'5');assert.deepEqual(w.getStyle('B1'),{});
});
test('REGRESSION: sorting formula rows recomputes references, styles and exact undo state',()=>{
  const w=create({A1:'Product',B1:'Qty',C1:'Value',A2:'Desk',B2:'12',C2:'=B2*5',A3:'Chair',B3:'3',C3:'=B3*5',A4:'Table',B4:'9',C4:'=B4*5'});
  w.patchStyles(['A3'],{italic:true,background:'#ffbb33'});
  const original=JSON.stringify(w.toJSON());
  const plan=sortRangeTransaction(w,{r1:0,r2:3,c1:0,c2:2},1,false,true);
  w.applyTransaction(plan.edits,plan.styles);
  assert.deepEqual(['A2','A3','A4'].map(r=>w.getRaw(r)),['Chair','Table','Desk']);
  assert.deepEqual(['C2','C3','C4'].map(r=>w.getValue(r).value),[15,45,60]);
  assert.equal(w.getStyle('A2').italic,true);
  assert.ok(w.undo());assert.equal(JSON.stringify(w.toJSON()),original);
});
test('REGRESSION: cycle recovery removes #CYCLE! from old dependency graph',()=>{
  const w=create({A1:'=B1+1',B1:'=A1+1',C1:'=A1+5'});
  assert.equal(w.getValue('C1').error?.code,'#CYCLE!');
  const changed=w.setCell('B1','10');
  assert.deepEqual(new Set(changed.affected),new Set(['A1','B1','C1']));
  assert.equal(w.getValue('C1').value,16);
  assert.equal(w.getValue('A1').error,undefined);
});
test('REGRESSION: relative vs locked references preserve expected positions',()=>{
  assert.equal(shiftFormula('=SUM($A1,B$2,$C$3,D4)',3,2),'=SUM($A4,D$2,$C$3,F7)');
});
test('REGRESSION: restored multi-sheet document does not share caches or styles',()=>{
  const d=new SpreadsheetDocument(create({A1:'5',B1:'=A1+1'}));
  d.addSheet('Second').book.setCell('A1','50');
  const copy=SpreadsheetDocument.fromJSON(JSON.parse(JSON.stringify(d.toJSON())));
  copy.sheets[0].book.setCell('A1','500');
  assert.equal(copy.sheets[0].book.getValue('B1').value,501);
  assert.equal(copy.sheets[1].book.getValue('A1').value,50);
  assert.equal(d.sheets[0].book.getRaw('A1'),'5');
});
test('REGRESSION: XLSX export remains a valid Office ZIP with correct parts', async()=>{
  const d=new SpreadsheetDocument(create({A1:'Name',A2:'Screw',B2:'3',C2:'=B2*4'}));
  d.renameSheet(d.activeId,'Sales');d.addSheet('Summary').book.setCell('A1','=SUM(1,2)');
  const entries=await zipRead(exportXlsx(d));
  for(const name of ['[Content_Types].xml','_rels/.rels','xl/workbook.xml','xl/styles.xml','xl/worksheets/sheet1.xml','xl/worksheets/sheet2.xml'])assert.ok(entries.has(name),name);
  assert.match(entries.get('xl/worksheets/sheet1.xml'),/<f>B2\*4<\/f>/);
  assert.match(entries.get('xl/worksheets/sheet2.xml'),/<f>SUM\(1,2\)<\/f>/);
});
