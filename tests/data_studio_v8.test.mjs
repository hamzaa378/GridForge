import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook} from '../dist/engine/workbook.js';
import {DATA_TOOLS, TOOL_CATEGORIES, previewDataTool} from '../dist/engine/tools_v8.js';

const rect=(r1=0,r2=0,c1=0,c2=0)=>({r1,r2,c1,c2});
const demo=(entries)=>{const w=new Workbook(400,52);w.applyEdits(entries.map(([ref,raw])=>({ref,raw:String(raw)})),false);return w;};
const textCases={
 upper:['Hello World','HELLO WORLD'],lower:['HELLO World','hello world'],title:['hELLO wORLD','Hello World'],sentence:['hELLO WORLD','Hello world'],
 swapcase:['HeLLo','hEllO'],'reverse-text':['Abc','cbA'],trim:['  hi  ','hi'],'collapse-spaces':['  a \t b  ','a b'],'remove-spaces':[' a b \t c ','abc'],'clean-control':['hi\n\tbye','hibye'],
 'first-word':['The quick fox','The'],'last-word':['The quick fox','fox'],'before-space':['one two','one'],'after-space':['one two three','two three'],
 initials:['john f kennedy','JFK'],'first-three':['abcdef','abc'],'last-three':['abcdef','def'],'drop-first':['abcdef','bcdef'],'drop-last':['abcdef','abcde'],'word-count':['one two three','3'],
 'remove-digits':['AB12-C3','AB-C'],'digits-only':['AB12-C3','123'],'letters-only':['A2 B-ç','ABç'],'alnum-only':['A2 B-ç','A2Bç'],
 'remove-punct':['Hi, team!','Hi team'],'remove-newlines':['a\nb\rc','a b c'],'tabs-to-spaces':['a\tb','a b'],
 'unicode-normalize':['ＡＢＣ','ABC'],'strip-accents':['café','cafe'],'strip-leading-zeros':['000123','123']
};
const numericCases={
 negate:[16,-16],absolute:[-16,16],reciprocal:[16,1/16],square:[16,256],cube:[16,4096],sqrt:[16,4],sign:[-16,-1],increment:[16,17],decrement:[16,15],double:[16,32],
 round0:[12.3456,12],round1:[12.3456,12.3],round2:[12.3456,12.35],round3:[12.3456,12.346],ceil:[12.3456,13],floor:[12.3456,12],trunc:[-12.3456,-12],
 nearest5:[12,10],nearest10:[16,20],nearest100:[160,200],divide100:[16,.16],multiply100:[16,1600],divide1000:[16,.016],multiply1000:[16,16000],half:[16,8],triple:[16,48],plus10pct:[16,17.6],minus10pct:[16,14.4],
 'parse-percent':['12.5%',.125],'parse-currency':['$1,234.50',1234.5]
};
const fillCases={
 'fill-down':['A2','5'],'fill-up':['A1','3'],'fill-right':['B1','5'],'fill-left':['A1','7'],
 'series-one':['B2','4'],'series-zero':['B2','3'],'series-even':['B2','8'],'series-odd':['B2','7'],'series-fives':['B2','20'],'series-tens':['B2','40']
};
const rangeCases={
 'flip-vertical':['A1','3'],'flip-horizontal':['A1','2'],'transpose-square':['B1','3'],'rotate-right':['A1','3'],'rotate-left':['A1','2'],
 'shift-down':['A1','3'],'shift-up':['A1','3'],'shift-right':['A1','2'],'shift-left':['A1','2'],'clear-diagonal':['A1','']
};
const qualityCases={
 'blank-zero':[['A1',''],['A1','0']], 'blank-na':[['A1',''],['A1','N/A']], 'zero-blank':[['A1','0'],['A1','']],
 'negative-zero':[['A1','-4'],['A1','0']], 'positive-zero':[['A1','4'],['A1','0']],
 'error-clear':[['A1','=1/0'],['A1','']], 'dedupe-rows':[['A2','red'],['A2','']], 'dedupe-cells':[['B1','red'],['B1','']],
 'space-hyphen':[['A1','red blue'],['A1','red-blue']], 'hyphen-space':[['A1','red-blue'],['A1','red blue']]
};
const aggCases={
 'insert-sum':['SUM',6], 'insert-average':['AVERAGE',3],'insert-min':['MIN',2],'insert-max':['MAX',4],
 'insert-count':['COUNT',2],'insert-counta':['COUNTA',2],'insert-median':['MEDIAN',3], 'insert-product':['PRODUCT',8],
 'insert-stdev':['STDEV.S',Math.sqrt(2)],'insert-countblank':['COUNTBLANK',0]
};

test('v8 has exactly 100 distinct, searchable, categorized spreadsheet actions',()=>{
 assert.equal(DATA_TOOLS.length,100);
 assert.equal(new Set(DATA_TOOLS.map(x=>x.id)).size,100);
 assert.equal(TOOL_CATEGORIES.length,10);
 for(const cat of TOOL_CATEGORIES)assert.equal(DATA_TOOLS.filter(x=>x.category===cat).length,10,cat);
 assert.equal(Object.keys(textCases).length,30);
 assert.equal(Object.keys(numericCases).length,30);
 assert.equal(Object.keys(fillCases).length,10);
 assert.equal(Object.keys(rangeCases).length,10);
 assert.equal(Object.keys(qualityCases).length,10);
 assert.equal(Object.keys(aggCases).length,10);
 assert.deepEqual(Object.keys({...textCases,...numericCases,...fillCases,...rangeCases,...qualityCases,...aggCases}).sort(),DATA_TOOLS.map(x=>x.id).sort());
});
for(const [id,[before,after]] of Object.entries(textCases))test(`Data Studio: ${id}`,()=>{
 const book=demo([['A1',before]]),previous=book.getRaw('A1');
 const change=previewDataTool(book,rect(),id);assert.equal(book.getRaw('A1'),previous,'Preview must be non-mutating');
 assert.deepEqual(change.edits,[{ref:'A1',raw:after}],id);
 book.applyEdits(change.edits);assert.equal(book.getRaw('A1'),after);
 assert.equal(book.undo(),true);assert.equal(book.getRaw('A1'),before);
 assert.equal(book.redo(),true);assert.equal(book.getRaw('A1'),after);
});
for(const [id,[before,after]] of Object.entries(numericCases))test(`Data Studio: ${id}`,()=>{
 const book=demo([['A1',before]]),change=previewDataTool(book,rect(),id);
 assert.equal(change.edits.length,1,`${id}: expected one edit`);
 const observed=Number(change.edits[0].raw);
 assert.ok(Math.abs(observed-after)<1e-10*Math.max(1,Math.abs(after)),`${id}: wanted ${after}, got ${observed}`);
 book.applyEdits(change.edits);assert.equal(book.undo(),true);assert.equal(book.getRaw('A1'),String(before));
});
for(const [id,[ref,expected]] of Object.entries(fillCases))test(`Data Studio: ${id}`,()=>{
 const book=demo([['A1','5'],['B1','7'],['A2','3'],['B2','9']]);
 const result=previewDataTool(book,rect(0,1,0,1),id);
 book.applyEdits(result.edits);
 assert.equal(book.getRaw(ref),expected,id);
 assert.equal(book.undo(),true);assert.equal(book.getRaw('A1'),'5');
});
for(const [id,[ref,expected]] of Object.entries(rangeCases))test(`Data Studio: ${id}`,()=>{
 const book=demo([['A1','1'],['B1','2'],['A2','3'],['B2','4']]);
 const result=previewDataTool(book,rect(0,1,0,1),id);
 book.applyEdits(result.edits);assert.equal(book.getRaw(ref),expected);
 assert.equal(book.undo(),true);assert.equal(book.getRaw('B2'),'4');
});
for(const [id,[[ref,before],[target,after]]] of Object.entries(qualityCases))test(`Data Studio: ${id}`,()=>{
 const seed=id==='dedupe-rows'? [['A1','red'],['A2',before]] :id==='dedupe-cells'?[['A1','red'],['B1',before]]:[[ref,before]];
 const book=demo(seed);const b=(id==='dedupe-rows')?rect(0,1,0,0):id==='dedupe-cells'?rect(0,0,0,1):rect();
 const result=previewDataTool(book,b,id);assert.ok(result.edits.length>0,id);
 book.applyEdits(result.edits);assert.equal(book.getRaw(target),after);
 assert.equal(book.undo(),true);assert.equal(book.getRaw(target),before);
});
for(const [id,[functionName,expected]] of Object.entries(aggCases))test(`Data Studio: ${id}`,()=>{
 const book=demo([['A1','2'],['A2','4']]);const result=previewDataTool(book,rect(0,1,0,0),id);
 assert.deepEqual(result.edits,[{ref:'A3',raw:`=${functionName}(A1:A2)`}]);
 book.applyEdits(result.edits);const val=book.getValue('A3');assert.equal(val.error,undefined);
 assert.ok(typeof val.value==='number'&&Math.abs(val.value-expected)<1e-9,`${id}: ${val.value} != ${expected}`);
 assert.equal(book.undo(),true);assert.equal(book.getRaw('A3'),'');
});

test('v8 text and numeric cleanup do not rewrite formulas',()=>{
 const book=demo([['A1','=SUM(1,2)'],['B1','  tEsT  '],['C1','3']]);
 for(const id of [...Object.keys(textCases),...Object.keys(numericCases)]){
  const change=previewDataTool(book,rect(0,0,0,0),id);
  assert.equal(change.edits.length,0,`${id} modified a formula`);
  assert.equal(book.getRaw('A1'),'=SUM(1,2)');
 }
});
test('v8 fill correctly adjusts formula references and absolute row/column locks',()=>{
 const book=demo([['A1','=B1+$D$1'],['B1','3'],['D1','7']]);
 const changes=previewDataTool(book,rect(0,2,0,0),'fill-down');
 assert.deepEqual(changes.edits,[{ref:'A2',raw:'=B2+$D$1'},{ref:'A3',raw:'=B3+$D$1'}]);
 book.applyEdits(changes.edits);assert.equal(book.getValue('A2').value,7);
});
test('v8 range transformations adjust formulas when moving and preserve undo',()=>{
 const book=demo([['A1','=B1'],['B1','12'],['A2','3'],['B2','4']]);
 const changes=previewDataTool(book,rect(0,1,0,1),'flip-vertical');
 book.applyEdits(changes.edits);assert.equal(book.getRaw('A2'),'=B2');assert.equal(book.undo(),true);assert.equal(book.getRaw('A1'),'=B1');
});
test('v8 rejects invalid, oversized, unsupported and out-of-bounds requests without mutation',()=>{
 const book=demo([['A1','5']]);
 for(const b of [rect(-1,0,0,0),rect(0,400,0,0),rect(0,399,0,40)])assert.throws(()=>previewDataTool(book,b,'upper'));
 assert.throws(()=>previewDataTool(book,rect(),'does-not-exist'));
 assert.throws(()=>previewDataTool(book,rect(0,1,0,0),'rotate-right'),/square/i);
 assert.throws(()=>previewDataTool(book,rect(399,399,0,0),'insert-sum'),/no room/i);
 assert.equal(book.getRaw('A1'),'5');
});
test('v8 handles invalid numeric values without NaN or Infinity stored in cells',()=>{
 const book=demo([['A1','0'],['A2','-9'],['A3','garbage']]);
 for(const [id,b] of [['reciprocal',rect(0,0,0,0)],['sqrt',rect(1,1,0,0)],['parse-currency',rect(2,2,0,0)]]){
  const result=previewDataTool(book,b,id);assert.equal(result.edits.length,0);assert.equal(result.skipped,1);
 }
});
test('v8 quality cleanup never silently converts valid formula results to raw values',()=>{
 const book=demo([['A1','=SUM(1,2)'],['B1','3']]);
 for(const id of ['blank-zero','blank-na','zero-blank','negative-zero','positive-zero','dedupe-cells']){
  const result=previewDataTool(book,rect(0,0,0,0),id);
  assert.equal(result.edits.length,0,id);
 }
});
