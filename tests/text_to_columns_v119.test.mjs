import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook} from '../dist/engine/workbook.js';
import {splitDelimitedText,textToColumnsEdits} from '../dist/engine/text_to_columns.js';

test('v11.9 text splitting honors quoted delimiters and escaped quote characters',()=>{
 assert.deepEqual(splitDelimitedText('North,"New York, NY",Approved',','),['North','New York, NY','Approved']);
 assert.deepEqual(splitDelimitedText('A,"She said ""yes""",C',','),['A','She said "yes"','C']);
 assert.throws(()=>splitDelimitedText('A,"unfinished',','),/not closed/);
});

test('v11.9 text splitting preserves empty fields or merges repeated delimiters',()=>{
 assert.deepEqual(splitDelimitedText('a,,b,',','),['a','','b','']);
 assert.deepEqual(splitDelimitedText('a,,b,',',',true),['a','b','']);
 assert.deepEqual(splitDelimitedText('a  b',' ',true),['a','b']);
});

test('v11.9 text to columns creates a rectangular undoable edit block at the chosen destination',()=>{
 const book=new Workbook(20,8);book.applyEdits([{ref:'A1',raw:'Name,City,State'},{ref:'A2',raw:'Hamza,"New Delhi, India"'},{ref:'A3',raw:'Riya,Chennai,TN'}],false);
 const plan=textToColumnsEdits(book,'A1','A3','C1',',');
 assert.equal(plan.rows,3);assert.equal(plan.columns,3);assert.equal(plan.edits.length,9);
 book.applyEdits(plan.edits);
 assert.deepEqual(['C1','D1','E1','C2','D2','E2','C3','D3','E3'].map(ref=>book.getRaw(ref)),['Name','City','State','Hamza','New Delhi, India','','Riya','Chennai','TN']);
});

test('v11.9 text to columns rejects multi-column sources and destination overflow',()=>{
 const book=new Workbook(5,4);book.applyEdits([{ref:'A1',raw:'a,b,c'},{ref:'A2',raw:'d,e,f'}],false);
 assert.throws(()=>textToColumnsEdits(book,'A1','B1','A1',','),/single-column/);
 assert.throws(()=>textToColumnsEdits(book,'A1','A2','D1',','),/last worksheet column/);
 assert.throws(()=>textToColumnsEdits(book,'A1','A2','A5',','),/last worksheet row/);
});
