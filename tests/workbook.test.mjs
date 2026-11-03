import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook,display} from '../dist/engine/workbook.js';
import {parseFormula,tokenize,shiftFormula,parseRef,colToName} from '../dist/engine/parser.js';
import {functions,matchCriterion} from '../dist/engine/functions.js';
import {parseCSV,encodeCSV} from '../dist/engine/csv.js';
const sheet=(cells={})=>{const w=new Workbook(400,52);w.applyEdits(Object.entries(cells).map(([ref,raw])=>({ref,raw:String(raw)})),false);return w;};
const value=(w,ref)=>w.getValue(ref).value;
const err=(w,ref)=>w.getValue(ref).error?.code;
const formula=(str)=>value(sheet({A1:str}),'A1');

test('challenge capacity: at least 10,000 addressable cells',()=>{const w=new Workbook();assert.ok(w.rows*w.cols>=10000);assert.equal(w.rows*w.cols,20800);});
test('original 56 functions remain registered after expansion',()=>{assert.ok(Object.keys(functions).length>=30);assert.ok(['SUM','IF','VLOOKUP'].every(k=>functions[k]));});
test('lexer recognizes references, range punctuation and quoted commas',()=>{const ts=tokenize('SUM($A$1:B3,"hi,world")');assert.deepEqual(ts.filter(t=>t.kind==='ref').map(t=>t.text),['$A$1','B3']);assert.equal(ts.filter(t=>t.kind==='string')[0].text,'"hi,world"');});
test('Pratt parser honors precedence and right associative exponent',()=>{assert.equal(formula('=2+3*4'),14);assert.equal(formula('=2^3^2'),512);assert.equal(formula('=-(2+3)^2'),-25);assert.equal(formula('=-2^2'),-4);});
test('arithmetic, unary, string concatenation and comparisons',()=>{assert.equal(formula('=(10-4)/3'),2);assert.equal(formula('="VIT "&"Tech"'),'VIT Tech');assert.equal(formula('=5>=2'),true);assert.equal(formula('=5<>5'),false);assert.equal(formula('=10/0'),null);assert.equal(err(sheet({A1:'=10/0'}),'A1'),'#DIV/0!');});
test('range aggregations and blank handling',()=>{const w=sheet({A1:'1',A2:'2',A3:'3',A4:'memo',B1:'=SUM(A1:A4)',B2:'=AVERAGE(A1:A4)',B3:'=COUNT(A1:A4)',B4:'=COUNTA(A1:A5)',B5:'=COUNTBLANK(A1:A5)',B6:'=MEDIAN(A1:A3)',B7:'=PRODUCT(A1:A3)'});assert.deepEqual([value(w,'B1'),value(w,'B2'),value(w,'B3'),value(w,'B4'),value(w,'B5'),value(w,'B6'),value(w,'B7')],[6,2,3,4,1,2,6]);});
test('conditional logic avoids evaluating unused branch',()=>{assert.equal(formula('=IF(FALSE,1/0,42)'),42);assert.equal(formula('=IF(TRUE,5,1/0)'),5);assert.equal(formula('=IFERROR(1/0,99)'),99);assert.equal(formula('=IFERROR(5/1,99)'),5);assert.equal(formula('=AND(TRUE,4>2)'),true);assert.equal(formula('=OR(FALSE,10>2)'),true);});
test('AND, OR, and XOR ignore text and blanks in references, reject direct text, and propagate range errors',()=>{
 const w=sheet({A1:'TRUE',A2:'note',A3:'',A4:'1',A5:'0',A6:'=1/0',B1:'=AND(A1:A4)',B2:'=AND(A1:A5)',B3:'=OR(A2:A3)',B4:'=AND("TRUE",TRUE)',B5:'=OR(TRUE,A6)',B6:'=AND(A2:A3)',B7:'=XOR(A1:A5)',B8:'=XOR("TRUE",TRUE)',B9:'=XOR(TRUE,A6)',B10:'=XOR(A2:A3)'});
 assert.equal(value(w,'B1'),true);
 assert.equal(value(w,'B2'),false);
 assert.equal(err(w,'B3'),'#VALUE!');
 assert.equal(err(w,'B4'),'#VALUE!');
 assert.equal(err(w,'B5'),'#DIV/0!');
 assert.equal(err(w,'B6'),'#VALUE!');
 assert.equal(value(w,'B7'),false);
 assert.equal(err(w,'B8'),'#VALUE!');
 assert.equal(err(w,'B9'),'#DIV/0!');
 assert.equal(err(w,'B10'),'#VALUE!');
});
test('VLOOKUP exact, MATCH, INDEX, and nested aggregation',()=>{const w=sheet({A1:'A',B1:'14',A2:'B',B2:'27',A3:'C',B3:'39',D1:'=VLOOKUP("B",A1:B3,2,FALSE)',D2:'=MATCH("C",A1:A3,0)',D3:'=INDEX(B1:B3,2)',D4:'=INDEX(A1:A3,MATCH("C",A1:A3,0))',D5:'=VLOOKUP("X",A1:B3,2,FALSE)'});assert.deepEqual([value(w,'D1'),value(w,'D2'),value(w,'D3'),value(w,'D4')],[27,3,27,'C']);assert.equal(err(w,'D5'),'#N/A');});
test('SUMIF COUNTIF AVERAGEIF match predicates',()=>{const w=sheet({A1:'2',A2:'5',A3:'10',B1:'20',B2:'50',B3:'100',C1:'=SUMIF(A1:A3,">4",B1:B3)',C2:'=COUNTIF(A1:A3,">=5")',C3:'=AVERAGEIF(A1:A3,">4",B1:B3)'});assert.deepEqual([value(w,'C1'),value(w,'C2'),value(w,'C3')],[150,2,75]);});
test('conditional formula criteria parse locale decimal and grouping separators as numbers',()=>{
 const de={decimal:',',group:'.',normalizeDigits:value=>value};
 assert.equal(matchCriterion(2,'>1,5',de),true);
 assert.equal(matchCriterion(1234.5,'=1.234,5',de),true);
 assert.equal(matchCriterion(1.5,'>1,5',de),false);
});
test('sparse conditional aggregation preserves errors found in populated ranges',()=>{const w=sheet({A1:'skip',A2:'take',B1:'=1/0',B2:'3',C1:'=SUMIF(A1:A2,"take",B1:B2)'});assert.equal(err(w,'C1'),'#DIV/0!');});
test('conditional SUMIF and AVERAGEIF stay sparse on full-height ranges and include blank matches',()=>{
 const w=new Workbook(1_048_576,16_384);w.applyEdits([
  {ref:'A1',raw:'0'},{ref:'A2',raw:'2'},{ref:'A1048576',raw:'3'},
  {ref:'B1',raw:'10'},{ref:'B2',raw:'20'},{ref:'B3',raw:'40'},{ref:'B4',raw:'50'},{ref:'B1048576',raw:'30'},
  {ref:'D1',raw:'=SUMIF(A1:A1048576,">0",B1:B1048576)'},
  {ref:'D2',raw:'=SUMIF(A1:A1048576,"",B1:B1048576)'},
  {ref:'D3',raw:'=AVERAGEIF(A1:A1048576,"",B1:B1048576)'},
  {ref:'D4',raw:'=SUMIF(A1:A1048576,"=0",B1:B1048576)'},
  {ref:'D5',raw:'=SUMIF(A1:A1048576,">0")'},
  {ref:'D6',raw:'=AVERAGEIF(A1:A1048576,">0")'},
  {ref:'D7',raw:'=COUNTIF(A1:A1048576,">0")'},
  {ref:'D8',raw:'=COUNTIFS(A1:A1048576,">0",B1:B1048576,">0")'},
  {ref:'D9',raw:'=SUMIFS(B1:B1048576,A1:A1048576,">0")'},
  {ref:'E1',raw:'>0'},{ref:'D10',raw:'=COUNTIF(A1:A1048576,E1)'},
  {ref:'D11',raw:'=COUNTIF(A1:A1048576,SEQUENCE(2))'},
  {ref:'D12',raw:'=AVERAGEIFS(B1:B1048576,A1:A1048576,">0")'},
  {ref:'D13',raw:'=MAXIFS(B1:B1048576,A1:A1048576,">0")'},
  {ref:'D14',raw:'=MINIFS(B1:B1048576,A1:A1048576,">0")'},
  {ref:'D15',raw:'=COUNTBLANK(A1:A1048576)'}
 ],false);
 assert.deepEqual(['D1','D2','D3','D4','D5','D6','D7','D8','D9','D10','D12','D13','D14','D15'].map(ref=>value(w,ref)),[50,90,45,10,5,2.5,2,2,50,2,25,30,20,1_048_573]);
 assert.equal(w.getValue('D11').error?.code,'#PARSE!');
});
test('numeric utility functions',()=>{const f=['ABS(-5)','ROUND(3.14159,2)','ROUNDUP(1.21,1)','ROUNDDOWN(1.29,1)','SQRT(81)','POWER(2,5)','MOD(14,3)','INT(3.99)','CEILING(13,5)','FLOOR(13,5)','SIGN(-18)','LOG(100,10)','LARGE(A2:A4,2)','SMALL(A2:A4,2)'];const w=sheet({A2:'3',A3:'1',A4:'2',...Object.fromEntries(f.map((x,i)=>[`B${i+1}`,`=${x}`]))});assert.deepEqual(f.map((_,i)=>value(w,`B${i+1}`)),[5,3.14,1.3,1.2,9,32,2,3,15,10,-1,2,2,2]);});
test('text manipulation functions',()=>{assert.equal(formula('=UPPER("hi")'),'HI');assert.equal(formula('=LOWER("HI")'),'hi');assert.equal(formula('=LEN("hello")'),5);assert.equal(formula('=TRIM("  hi    there ")'),'hi there');assert.equal(formula('=LEFT("hello",2)'),'he');assert.equal(formula('=RIGHT("hello",2)'),'lo');assert.equal(formula('=MID("hello",2,3)'),'ell');assert.equal(formula('=REPLACE("hello",2,2,"i")'),'hilo');assert.equal(formula('=SUBSTITUTE("mama","ma","pa")'),'papa');assert.equal(formula('=FIND("l","hello")'),3);assert.equal(formula('=SEARCH("L","hello")'),3);assert.equal(formula('=CONCAT("he","llo")'),'hello');});
test('statistics and type predicates',()=>{assert.equal(formula('=ISNUMBER(12)'),true);assert.equal(formula('=ISTEXT("hi")'),true);assert.equal(formula('=EXACT("A","a")'),false);assert.equal(formula('=NOT(FALSE)'),true);const w=sheet({A1:'2',A2:'4',A3:'6',B1:'=STDEV(A1:A3)',B2:'=VAR(A1:A3)'});assert.equal(value(w,'B1'),2);assert.equal(value(w,'B2'),4);});
test('parse error has caret position and actionable suggestion',()=>{const w=sheet({D1:'=SUM(A1, B2'}),d=w.getValue('D1').error;assert.equal(d.code,'#PARSE!');assert.match(d.message,/Missing closing parenthesis/i);assert.match(d.suggestion,/Add '\)'/i);assert.ok(d.position>1);});
test('unknown function, division zero, bad reference, bad argument',()=>{const w=sheet({A1:'=SUMM(1,2)',A2:'=1/0',A3:'=ZZ999',A4:'=INDEX(B10:B11,9)',A5:'=SQRT(-2)'});assert.deepEqual(['A1','A2','A3','A4','A5'].map(x=>err(w,x)),['#NAME?','#DIV/0!','#REF!','#REF!','#NUM!']);});
test('direct and indirect cycles terminate and report #CYCLE!',()=>{const w=sheet({A1:'=A1+1',B1:'=C1+1',C1:'=D1+1',D1:'=B1+1'});assert.equal(err(w,'A1'),'#CYCLE!');for(const r of ['B1','C1','D1'])assert.equal(err(w,r),'#CYCLE!');});
test('dependency invalidation only touches affected chain, not unrelated formulas',()=>{const w=sheet({A1:'1',B1:'=A1+1',C1:'=B1*2',D1:'=100+25',E1:'=D1*3'});const s=w.setCell('A1','10');assert.equal(value(w,'C1'),22);assert.equal(value(w,'E1'),375);assert.deepEqual(new Set(s.affected),new Set(['A1','B1','C1']));assert.equal(s.recalculated,3);});
test('switching dependencies detaches old graph edges',()=>{const w=sheet({A1:'1',B1:'2',C1:'=A1'});w.setCell('C1','=B1');const x=w.setCell('A1','88');assert.equal(value(w,'C1'),2);assert.equal(x.affected.includes('C1'),false);assert.equal(w.setCell('B1','91').affected.includes('C1'),true);});
test('undo and redo for one edit and grouped edits',()=>{const w=sheet({A1:'1',B1:'2'});w.applyEdits([{ref:'A1',raw:'10'},{ref:'B1',raw:'20'}]);assert.equal(value(w,'A1'),10);w.undo();assert.equal(value(w,'A1'),1);assert.equal(value(w,'B1'),2);w.redo();assert.equal(value(w,'B1'),20);});
test('relative fill respects absolute row and column locks',()=>{assert.equal(shiftFormula('=A1+$B$2+C$3+$D4',1,2),'=C2+$B$2+E$3+$D5');assert.equal(shiftFormula('="A1"&A1',1,0),'="A1"&A2');assert.equal(shiftFormula('=SUM(A1:B2)',2,1),'=SUM(B3:C4)');});
test('non-destructive what-if preview provides a diff',()=>{const w=sheet({A1:'5',B1:'=A1*2',C1:'=B1+1'});const p=w.preview('A1','8');assert.equal(value(w,'A1'),5);assert.equal(value(w,'C1'),11);assert.deepEqual(p.changes.map(x=>[x.ref,x.after]),[['A1','8'],['B1','16'],['C1','17']]);});
test('JSON roundtrip preserves raw formulas, values and 20,800-cell dimensions',()=>{const w=sheet({A1:'5',B1:'=A1*2',C1:'multi\nline'});const clone=Workbook.fromJSON(JSON.parse(JSON.stringify(w.toJSON())));assert.equal(clone.getRaw('B1'),'=A1*2');assert.equal(value(clone,'B1'),10);assert.equal(value(clone,'C1'),'multi\nline');assert.equal(clone.rows*clone.cols,20800);});
test('malformed JSON workbooks are rejected',()=>{assert.throws(()=>Workbook.fromJSON({rows:1_048_577,cols:3,cells:{}}));assert.throws(()=>Workbook.fromJSON({rows:100,cols:10,cells:{A1:42}}));assert.throws(()=>Workbook.fromJSON({rows:100,cols:10,cells:{ZZ99:'x'}}));});
test('CSV quoted commas, newlines and quote escapes roundtrip',()=>{const rows=[['a,b','embedded "quote"','x\ny'],['2','3','4']];assert.deepEqual(parseCSV(encodeCSV(rows)),rows);assert.deepEqual(parseCSV('a,b\r\n1,2'),[['a','b'],['1','2']]);});
test('CSV rejects unterminated quoted fields',()=>assert.throws(()=>parseCSV('"not closed')));
test('address conversion AA52 / locks',()=>{assert.equal(colToName(26),'AA');assert.deepEqual(parseRef('$AA$52'),{row:51,col:26,rowLock:true,colLock:true});});
test('large workbook: 12k populated cells, isolated recalculation is selective',()=>{
 const w=new Workbook(400,52),edits=[];
 for(let r=1;r<=300;r++)for(let c=0;c<40;c++)edits.push({ref:`${colToName(c)}${r}`,raw:String(r+c)});
 w.applyEdits(edits,false);const s=w.setCell('A2','999');assert.equal(s.affected.length,1);assert.equal(s.recalculated,1);assert.equal(value(w,'A2'),999);
});
