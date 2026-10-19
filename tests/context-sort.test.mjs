import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook} from '../dist/engine/workbook.js';
import {sortRangeTransaction} from '../dist/engine/operations.js';
import {snapshotCells,pasteSnapshot} from '../dist/engine/clipboard.js';
const make=(obj)=>{const w=new Workbook(100,25);w.applyEdits(Object.entries(obj).map(([ref,raw])=>({ref,raw})),false);return w;};
const selected=(r1,r2,c1,c2)=>({r1,r2,c1,c2});
test('custom sort by any selected column keeps complete rows and moves formatting',()=>{
 const w=make({Q1:'Product',R1:'Qty',S1:'Price',Q2:'Keyboard',R2:'50',S2:'1500',Q3:'Mouse',R3:'10',S3:'500',Q4:'Monitor',R4:'80',S4:'12000',Q5:'Webcam',R5:'25',S5:'2500'});
 w.patchStyles(['Q3'],{bold:true,color:'#abcdef'});
 const plan=sortRangeTransaction(w,selected(0,4,16,18),17,false,true);
 w.applyTransaction(plan.edits,plan.styles);
 assert.deepEqual([w.getRaw('R2'),w.getRaw('R3'),w.getRaw('R4'),w.getRaw('R5')],['10','25','50','80']);
 assert.deepEqual([w.getRaw('Q2'),w.getRaw('Q3'),w.getRaw('Q4'),w.getRaw('Q5')],['Mouse','Webcam','Keyboard','Monitor']);
 assert.equal(w.getRaw('Q1'),'Product');assert.equal(w.getStyle('Q2').bold,true);
 w.undo();assert.equal(w.getRaw('Q2'),'Keyboard');assert.deepEqual(w.getStyle('Q2'),{});
 assert.equal(w.getStyle('Q3').color,'#abcdef');w.redo();assert.equal(w.getStyle('Q2').color,'#abcdef');
});
test('sort changes only selected columns, and stable duplicates preserve input order',()=>{
 const w=make({A1:'10',B1:'first',C1:'outside1',A2:'2',B2:'second',C2:'outside2',A3:'2',B3:'third',C3:'outside3'});
 const tx=sortRangeTransaction(w,selected(0,2,0,1),0);w.applyTransaction(tx.edits,tx.styles);
 assert.deepEqual([w.getRaw('B1'),w.getRaw('B2'),w.getRaw('B3')],['second','third','first']);
 assert.deepEqual([w.getRaw('C1'),w.getRaw('C2'),w.getRaw('C3')],['outside1','outside2','outside3']);
});
test('copy/paste adjusts relative formulas, includes styles and can undo',()=>{
 const w=make({A1:'3',B1:'=A1*2'});w.patchStyles(['B1'],{bold:true});
 const snap=snapshotCells(w,selected(0,0,1,1),'s1');
 const plan=pasteSnapshot(w,snap,{row:1,col:1},'s1');w.applyTransaction(plan.edits,plan.styles);
 assert.equal(w.getRaw('B2'),'=A2*2');assert.equal(w.getStyle('B2').bold,true);
 w.undo();assert.equal(w.getRaw('B2'),'');assert.deepEqual(w.getStyle('B2'),{});
});
test('cut then paste moves cell and style atomically; overlapping moves are safe',()=>{
 const w=make({A1:'One',B1:'Two',C1:'Three'});w.patchStyles(['A1'],{italic:true});
 const snap=snapshotCells(w,selected(0,0,0,1),'s1',true);
 const plan=pasteSnapshot(w,snap,{row:0,col:1},'s1');w.applyTransaction(plan.edits,plan.styles);
 assert.deepEqual([w.getRaw('A1'),w.getRaw('B1'),w.getRaw('C1')],['','One','Two']);
 assert.equal(w.getStyle('B1').italic,true);
 w.undo();assert.deepEqual([w.getRaw('A1'),w.getRaw('B1'),w.getRaw('C1')],['One','Two','Three']);
 assert.equal(w.getStyle('A1').italic,true);
});
test('cut from a different worksheet becomes copy, never clears source',()=>{
 const w=make({A1:'Keep'});const snap=snapshotCells(w,selected(0,0,0,0),'s1',true);
 const other=make({});const plan=pasteSnapshot(other,snap,{row:1,col:1},'s2');
 assert.equal(plan.moved,false);other.applyTransaction(plan.edits,plan.styles);
 assert.equal(w.getRaw('A1'),'Keep');assert.equal(other.getRaw('B2'),'Keep');
});
test('paste values preserves leading equals as literal text and prevents overflow',()=>{
 const w=make({A1:"'=SUM(A2)",A2:'4'});
 const snap=snapshotCells(w,selected(0,0,0,0),'s1');
 const tx=pasteSnapshot(w,snap,{row:1,col:1},'s1',true);w.applyTransaction(tx.edits,tx.styles);
 assert.equal(w.getValue('B2').value,'=SUM(A2)');
 assert.throws(()=>pasteSnapshot(w,snap,{row:100,col:1},'s1'),/outside/);
});
test('sorting with zero data rows returns empty transaction',()=>{
 const w=make({A1:'Header',A2:'only row'});
 const plan=sortRangeTransaction(w,selected(0,1,0,0),0,false,true);
 assert.deepEqual(plan,{edits:[],styles:[]});
});
