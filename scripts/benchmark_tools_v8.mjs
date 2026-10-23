#!/usr/bin/env node
/** Deterministic 10,000-cell Data Studio stress and undo integrity check. */
import assert from 'node:assert/strict';
import {performance} from 'node:perf_hooks';
import {Workbook} from '../dist/engine/workbook.js';
import {previewDataTool,DATA_TOOLS} from '../dist/engine/tools_v8.js';
const book=new Workbook(400,52);
const edits=[];
for(let r=0;r<200;r++)for(let c=0;c<50;c++){
 let n=c+1,name='';while(n){n--;name=String.fromCharCode(65+n%26)+name;n=Math.floor(n/26);}
 edits.push({ref:`${name}${r+1}`,raw:String(r*50+c+1)});
}
book.applyEdits([...edits,{ref:'AZ400',raw:'UNRELATED'}],false);
const bounds={r1:0,r2:199,c1:0,c2:49},start=performance.now();
const preview=previewDataTool(book,bounds,'double');
const previewMs=performance.now()-start;
assert.equal(preview.edits.length,10000);
assert.equal(book.getRaw('A1'),'1','Preview must not alter workbook');
const begin=performance.now();book.applyEdits(preview.edits);const applyMs=performance.now()-begin;
assert.equal(book.getRaw('A1'),'2');assert.equal(book.getRaw('AX200'),'20000');
assert.equal(book.getRaw('AZ400'),'UNRELATED');
assert.equal(book.undo(),true);assert.equal(book.getRaw('A1'),'1');
assert.equal(book.getRaw('AX200'),'10000');assert.equal(book.getRaw('AZ400'),'UNRELATED');
assert.equal(DATA_TOOLS.length,100);
assert.throws(()=>previewDataTool(book,{r1:0,r2:399,c1:0,c2:49},'double'),/12,000/);
console.log(JSON.stringify({features:DATA_TOOLS.length,selectedCells:10000,edited:preview.edits.length,previewMs:+previewMs.toFixed(2),applyMs:+applyMs.toFixed(2),undoVerified:true,unrelatedPreserved:true,oversizeRejected:true},null,2));
