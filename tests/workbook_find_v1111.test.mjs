import test from 'node:test';
import assert from 'node:assert/strict';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {findAcrossSheets} from '../dist/engine/operations.js';

test('v11.11 Find All returns worksheet-qualified matches across visible sheets',()=>{
 const doc=new SpreadsheetDocument(),first=doc.active,second=doc.addSheet('Second'),hidden=doc.addSheet('Hidden');
 first.book.applyEdits([{ref:'A1',raw:'Target'},{ref:'A2',raw:'=Target'}],false);
 second.book.applyEdits([{ref:'C3',raw:'target row'}],false);
 hidden.book.applyEdits([{ref:'D4',raw:'Target'}],false);doc.hideSheet(hidden.id);
 const matches=findAcrossSheets(doc.sheets,'target',false);
 assert.deepEqual(matches.map(match=>[match.sheetName,match.ref]),[['Sheet 1','A1'],['Sheet 1','A2'],['Second','C3']]);
 assert.equal(findAcrossSheets(doc.sheets,'target',false,false).length,2);
});
