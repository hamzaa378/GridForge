import fs from 'node:fs';
const cases=[];
function add(id,formula,{inputs={},category='core',expected,expectedSpill,kind='scalar',note=''}={}){
  if(cases.some(c=>c.id===id))throw Error('Duplicate id '+id);
  cases.push({id,formula,inputs,category,kind,...(expected===undefined?{}:{referenceExpected:expected}),...(expectedSpill===undefined?{}:{expectedSpill}),...(note?{note}:{})});
}
// Expected values below are fixed reference examples, not a claim that Excel was run locally.
const math=[
['round-positive-half','=ROUND(1.5,0)',2],['round-negative-half','=ROUND(-1.5,0)',-2],
['round-negative-two-half','=ROUND(-2.5,0)',-3],['round-negative-decimal','=ROUND(-1.475,2)',-1.48],
['round-positive-one','=ROUND(2.15,1)',2.2],['round-negative-digits','=ROUND(626.3,-3)',1000],
['roundup-negative','=ROUNDUP(-3.14159,1)',-3.2],['rounddown-negative','=ROUNDDOWN(-3.99,1)',-3.9],
['abs-negative','=ABS(-19)',19],['int-negative','=INT(-3.2)',-4],
['sqrt-perfect','=SQRT(81)',9],['power-integer','=POWER(2,8)',256],
['mod-negative-divisor','=MOD(3,-2)',-1],['mod-negative-numerator','=MOD(-3,2)',1],
['sign-negative','=SIGN(-4)',-1],['factorial-five','=FACT(5)',120],
['floor-positive','=FLOOR(13,5)',10],['ceiling-positive','=CEILING(13,5)',15],
['logten','=LOG10(100)',2],['ln-one','=LN(1)',0],
['sin-zero','=SIN(0)',0],['cos-zero','=COS(0)',1],
['min-simple','=MIN(1,4,2)',1],['max-simple','=MAX(1,4,2)',4],
['average-three','=AVERAGE(1,2,3)',2],['product-three','=PRODUCT(2,3,4)',24],
['sum-three','=SUM(2,3,4)',9],['count-literal','=COUNT(3,5)',2],
['sum-coercion-text','=SUM("3",4)',7],['sum-coercion-boolean','=SUM(TRUE,2)',3],
['count-coercion-text','=COUNT(3,"4")',2]
]; for(const [id,formula,expected] of math)add(id,formula,{category:'math',expected});
const ops=[
['add','=2+3',5],['precedence','=2+3*4',14],['parens','=(2+3)*4',20],['divide','=9/3',3],
['exponent','=2^3',8],['concatenate','="ab"&"cd"','abcd'],['equality','=3=3',true],
['less-than','=2<3',true],['boolean-if','=IF(TRUE,10,20)',10],
['if-short-circuit','=IF(TRUE,100,1/0)',100],['iferror-short-circuit','=IFERROR(10,1/0)',10],
['iferror-handles','=IFERROR(1/0,9)',9],['ifna-handles','=IFNA(NA(),7)',7],
['logic-and','=AND(TRUE,TRUE,FALSE)',false],['logic-or','=OR(FALSE,TRUE)',true],
['logic-not','=NOT(TRUE)',false],['len-ascii','=LEN("hello")',5],
['left-text','=LEFT("hello",2)','he'],['right-text','=RIGHT("hello",2)','lo'],
['mid-text','=MID("hello",2,3)','ell'],['upper','=UPPER("MiXeD")','MIXED'],
['lower','=LOWER("MiXeD")','mixed'],['trim','=TRIM("  a   b  ")','a b'],
['substitute','=SUBSTITUTE("abab","ab","x")','xx'],
['find','=FIND("lo","hello")',4],['exact-case','=EXACT("A","a")',false],
['textjoin','=TEXTJOIN(",",TRUE,"a","b","c")','a,b,c'],
['sumproduct','=SUMPRODUCT(A1:A3,B1:B3)',32]
];for(const [id,formula,expected] of ops)add(id,formula,{category:'logic-text',expected,inputs:id==='sumproduct'?{A1:1,A2:2,A3:3,B1:4,B2:5,B3:6}:{}});
const refCases=[
 ['range-sum','=SUM(A1:A4)',{A1:2,A2:3,A3:4,A4:5},14],
 ['range-blank-count','=COUNTBLANK(A1:A4)',{A1:2,A3:4},2],
 ['range-count-numeric','=COUNT(A1:A4)',{A1:2,A2:'foo',A3:4,A4:5},3],
 ['range-text-ignored','=SUM(A1:A3)',{A1:2,A2:{text:'3'},A3:4},6],
 ['range-average','=AVERAGE(A1:A4)',{A1:2,A2:4,A3:6,A4:8},5],
 ['range-sample-stdev','=STDEV.S(A1:A3)',{A1:1,A2:2,A3:3},1],
 ['range-var-s','=VAR.S(A1:A3)',{A1:1,A2:2,A3:3},1],
 ['range-var-p','=VAR.P(A1:A3)',{A1:1,A2:2,A3:3},2/3],
 ['sumif-north','=SUMIF(A1:A4,"North",B1:B4)',{A1:'North',A2:'South',A3:'North',A4:'East',B1:4,B2:7,B3:5,B4:9},9],
 ['countif-ten','=COUNTIF(A1:A4,">10")',{A1:7,A2:11,A3:12,A4:4},2],
 ['vlookup-exact','=VLOOKUP("Mouse",A1:B3,2,FALSE)',{A1:'Mouse',B1:12,A2:'Keyboard',B2:14,A3:'Monitor',B3:99},12],
 ['index-lookup','=INDEX(A1:B3,2,2)',{A1:1,B1:2,A2:3,B2:4,A3:5,B3:6},4],
 ['match-exact','=MATCH("B",A1:A3,0)',{A1:'A',A2:'B',A3:'C'},2],
 ['absolute-ref','=A1+$B$1',{A1:2,B1:3},5],
 ['formula-dependency','=SUM(A1:A3)*2',{A1:1,A2:2,A3:3},12],
 ['countif-wildcard','=COUNTIF(A1:A3,"A*")',{A1:'Apple',A2:'Apricot',A3:'Banana'},2],
 ['sumifs-multi','=SUMIFS(C1:C4,A1:A4,"North",B1:B4,">5")',{A1:'North',A2:'South',A3:'North',A4:'North',B1:3,B2:6,B3:7,B4:9,C1:2,C2:4,C3:6,C4:8},14]
];for(const [id,formula,inputs,expected] of refCases)add(id,formula,{inputs,category:'references',expected});
const dates=[
 ['date-march-1900','=DATE(1900,3,1)',61],
 ['date-2020','=DATE(2020,1,1)',43831],
 ['year-from-date','=YEAR(DATE(2024,2,1))',2024],
 ['month-from-date','=MONTH(DATE(2024,2,1))',2],
 ['day-from-date','=DAY(DATE(2024,2,29))',29],
 ['days-simple','=DAYS(DATE(2024,1,10),DATE(2024,1,1))',9],
 ['edate-month','=EDATE(DATE(2024,1,31),1)',45351]
];for(const [id,formula,expected] of dates)add(id,formula,{category:'dates',expected});
const errors=[['division-zero','=1/0','#DIV/0!'],['negative-sqrt','=SQRT(-1)','#NUM!'],['unknown-function','=NOTAFUNCTION(5)','#NAME?'],['na','=NA()','#N/A'],['bad-ref','=INDEX(A1:B2,10,1)','#REF!']];
for(const [id,formula,expected] of errors)add(id,formula,{category:'errors',expected});
for(const x of [0,1,2,3,5,7,13,17,29,37,43,71,97,127,199,255]) {
 add(`math-sum-${x}`,`=SUM(${x},${x+2},${x+4})`,{expected:3*x+6,category:'generated-math'});
 add(`math-divide-${x}`,`=(${x+5}*6)/3`,{expected:(x+5)*2,category:'generated-math'});
 add(`math-round-neg-${x}`,`=ROUND(-${x}.5,0)`,{expected:-(x+1),category:'generated-math'});
}
// Array results are checked at the anchor and across the spill range.
for(const [id,formula,expectedSpill] of [
 ['dynamic-sequence-spill','=SEQUENCE(2,2)',[[1,2],[3,4]]],
 ['dynamic-filter-spill','=FILTER(A1:A3,B1:B3)',[[1],[1]]],
 ['dynamic-unique-spill','=UNIQUE(A1:A3)',[[1],[2]]],
 ['dynamic-textsplit-spill','=TEXTSPLIT("A,B",",")',[["A","B"]]]]){
 add(id,formula,{category:'dynamic-arrays',kind:'spill',expected:expectedSpill[0][0],expectedSpill,inputs:{A1:1,A2:2,A3:1,B1:1,B2:0,B3:1},note:'GridForge spills the result into adjacent cells; a Microsoft Excel comparison is still required for parity evidence.'});
}
// Avoid volatile functions and network/locale-specific functions in the scalar agreement denominator.
const output={schema:'gridforge-excel-parity/1',description:'Deterministic cross-engine formula probes; Excel execution required for true differential results.',comparisonCell:'K5',cases};
fs.writeFileSync(new URL('../tests/excel_parity_cases.json',import.meta.url),JSON.stringify(output,null,2)+'\n');
console.log(`${cases.length} cases generated; ${cases.filter(c=>c.kind==='scalar').length} scalar, ${cases.filter(c=>c.kind==='spill').length} spill cases`);
