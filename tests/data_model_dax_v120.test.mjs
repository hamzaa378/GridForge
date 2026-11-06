import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook} from '../dist/engine/workbook.js';
import {buildDataModel,evaluateDaxMeasure,validateDataModel} from '../dist/engine/data_model.js';
import {excelDateSerial} from '../dist/engine/date_serial.js';

function runtime(){
 const book=new Workbook(12,5);book.applyEdits([
  {ref:'A1',raw:'Region'},{ref:'B1',raw:'Amount'},
  {ref:'A2',raw:'East'},{ref:'B2',raw:''},
  {ref:'A3',raw:'West'},{ref:'B3',raw:'5'},
  {ref:'A4',raw:'East'},{ref:'B4',raw:'7'},
  {ref:'A5',raw:''},{ref:'B5',raw:'0'},
 ],false);
 return buildDataModel([{name:'Data',book}],{tables:[{name:'Sales',sheet:'Data',range:'A1:B5'}],relationships:[],measures:[]});
}

test('local DAX counts blanks in DISTINCTCOUNT and returns blank for empty table aggregates',()=>{
 const model=runtime();
 assert.equal(evaluateDaxMeasure(model,'DISTINCTCOUNT(Sales[Region])'),3);
 assert.equal(evaluateDaxMeasure(model,'DISTINCTCOUNTNOBLANK(Sales[Region])'),2);
 assert.equal(evaluateDaxMeasure(model,'COUNTBLANK(Sales[Region])'),1);
 assert.equal(evaluateDaxMeasure(model,'COUNTX(Sales, Sales[Region])'),0);
 assert.equal(evaluateDaxMeasure(model,'COUNTAX(Sales, Sales[Region])'),3);
 assert.equal(evaluateDaxMeasure(model,'COUNTAX(Sales, Sales[Region] & "!")'),4);
 assert.equal(evaluateDaxMeasure(model,'SUMX(Sales, IF(Sales[Region] & "" = "West", Sales[Amount], 0))'),5);
 assert.equal(evaluateDaxMeasure(model,'IF(("A" & "B") = "AB", 1, 0)'),1);
 assert.equal(evaluateDaxMeasure(model,'SUMX(Sales, Sales[Amount])'),12);
 assert.equal(evaluateDaxMeasure(model,'MEDIAN(Sales[Amount])'),5);
 assert.equal(evaluateDaxMeasure(model,'MEDIANX(Sales, COALESCE(Sales[Amount], 0))'),2.5);
 assert.equal(evaluateDaxMeasure(model,'PRODUCT(Sales[Amount])'),0);
 assert.equal(evaluateDaxMeasure(model,'PRODUCTX(Sales, IF(Sales[Amount] > 0, Sales[Amount], 2))'),140);
 assert.equal(evaluateDaxMeasure(model,'COUNTROWS(Sales)',[{table:'Sales',column:'Region',values:['Nowhere']}]),null);
 assert.equal(evaluateDaxMeasure(model,'DISTINCTCOUNTNOBLANK(Sales[Region])',[{table:'Sales',column:'Region',values:['Nowhere']}]),null);
 assert.equal(evaluateDaxMeasure(model,'COUNTROWS(VALUES(Sales[Region]))'),3);
 assert.equal(evaluateDaxMeasure(model,'COUNTROWS(DISTINCT(Sales[Amount]))'),4);
 assert.equal(evaluateDaxMeasure(model,'SUMX(VALUES(Sales[Amount]), Sales[Amount])'),12);
 assert.equal(evaluateDaxMeasure(model,'SUMX(ADDCOLUMNS(Sales, "Double", COALESCE(Sales[Amount], 0) * 2), Sales[Double])'),24);
 assert.equal(evaluateDaxMeasure(model,'COUNTROWS(FILTER(ADDCOLUMNS(Sales, "Double", COALESCE(Sales[Amount], 0) * 2), Sales[Double] > 10))'),1);
 assert.throws(()=>evaluateDaxMeasure(model,'COUNTROWS(ADDCOLUMNS(Sales, "Region", Sales[Amount]))'),/unique text column names/);
 assert.equal(evaluateDaxMeasure(model,'SUMX(SELECTCOLUMNS(Sales, "Value", COALESCE(Sales[Amount], 0)), Sales[Value])'),12);
 assert.equal(evaluateDaxMeasure(model,'COUNTROWS(SELECTCOLUMNS(Sales, "Label", Sales[Region]))'),4);
 assert.equal(evaluateDaxMeasure(model,'COUNTROWS(SUMMARIZE(Sales, Sales[Region]))'),3);
 assert.equal(evaluateDaxMeasure(model,'COUNTROWS(SUMMARIZE(Sales, Sales[Region], Sales[Amount]))'),4);
 assert.equal(evaluateDaxMeasure(model,'COUNTROWS(SUMMARIZE(Sales, Sales[Region]))',[{table:'Sales',column:'Region',values:['East']}]),1);
 assert.throws(()=>evaluateDaxMeasure(model,'COUNTROWS(VALUES(Sales[Missing]))'),/Unknown model column/);
 assert.equal(evaluateDaxMeasure(model,'COUNTROWS(ALL(Sales))',[{table:'Sales',column:'Region',values:['Nowhere']}]),4);
 assert.equal(evaluateDaxMeasure(model,'COUNTROWS(ALL(Sales[Region]))',[{table:'Sales',column:'Region',values:['Nowhere']}]),3);
 assert.equal(evaluateDaxMeasure(model,'SUMX(ALL(Sales), Sales[Amount])',[{table:'Sales',column:'Region',values:['Nowhere']}]),12);
 assert.equal(evaluateDaxMeasure(model,'SUM(Sales[Amount])',[{table:'Sales',column:'Region',values:['Nowhere']}]),null);
});

test('local DAX supports selected-value and direct-filter scalar helpers',()=>{
 const model=runtime();
 assert.equal(evaluateDaxMeasure(model,'SELECTEDVALUE(Sales[Region], "multiple")'),'multiple');
 assert.equal(evaluateDaxMeasure(model,'SELECTEDVALUE(Sales[Region], "multiple")',[{table:'Sales',column:'Region',values:['East']}]),'East');
 assert.equal(evaluateDaxMeasure(model,'HASONEVALUE(Sales[Region])',[{table:'Sales',column:'Region',values:['East']}]),true);
 assert.equal(evaluateDaxMeasure(model,'ISFILTERED(Sales[Region])',[{table:'Sales',column:'Region',values:['East']}]),true);
 assert.equal(evaluateDaxMeasure(model,'CALCULATE(SELECTEDVALUE(Sales[Region], "multiple"), Sales[Region] = "West")'),'West');
 assert.equal(evaluateDaxMeasure(model,'ISBLANK(BLANK())'),true);
});

test('local DAX FILTER provides bounded row-filtered tables to iterators and COUNTROWS',()=>{
 const model=runtime();
 assert.equal(evaluateDaxMeasure(model,'COUNTROWS(FILTER(Sales, Sales[Amount] > 5))'),1);
 assert.equal(evaluateDaxMeasure(model,'SUMX(FILTER(Sales, Sales[Region] = "East"), Sales[Amount])'),7);
 assert.equal(evaluateDaxMeasure(model,'SUMX(FILTER(FILTER(Sales, Sales[Amount] > 0), Sales[Region] = "East"), Sales[Amount])'),7);
 assert.equal(evaluateDaxMeasure(model,'COUNTROWS(FILTER(Sales, Sales[Amount] > 100))'),null);
 assert.equal(evaluateDaxMeasure(model,'COUNTROWS(FILTER(Sales, Sales[Region] = "East"))',[{table:'Sales',column:'Region',values:['West']}]),null);
 assert.throws(()=>evaluateDaxMeasure(model,'COUNTROWS(FILTER(Sales, Sales[Amount]))'),/must return TRUE or FALSE/);
 assert.throws(()=>evaluateDaxMeasure(model,'COUNTROWS(FILTER(Sales, Sales[Missing] > 1))'),/Unknown model column/);
});

test('local DAX rejects excessive expression nesting before evaluation',()=>{
 const nested='('.repeat(65)+'1'+')'.repeat(65);
 assert.throws(()=>validateDataModel({tables:[],relationships:[],measures:[{name:'Deep',expression:nested}]}),/nesting is limited/);
 assert.throws(()=>validateDataModel({tables:[],relationships:[],measures:[{name:'Deep',expression:'-'.repeat(65)+'1'}]}),/unary operator nesting/);
 assert.throws(()=>validateDataModel({tables:[],relationships:[],measures:[{name:'Deep',expression:'NOT '.repeat(65)+'TRUE'}]}),/unary operator nesting/);
 assert.doesNotThrow(()=>validateDataModel({tables:[],relationships:[],measures:[{name:'Text',expression:'IF(TRUE, "'+'('.repeat(100)+'", "")'}]}));
});

test('local model propagates date filters across relationship chains and resolves RELATED paths',()=>{
 const dates=new Workbook(8,3),calendar=new Workbook(8,3),sales=new Workbook(8,3);
 dates.applyEdits([{ref:'A1',raw:'DateID'},{ref:'B1',raw:'Year'},{ref:'A2',raw:'1'},{ref:'B2',raw:'2024'},{ref:'A3',raw:'2'},{ref:'B3',raw:'2025'}],false);
 calendar.applyEdits([{ref:'A1',raw:'CalendarID'},{ref:'B1',raw:'DateID'},{ref:'A2',raw:'100'},{ref:'B2',raw:'1'},{ref:'A3',raw:'101'},{ref:'B3',raw:'2'}],false);
 sales.applyEdits([{ref:'A1',raw:'CalendarID'},{ref:'B1',raw:'Amount'},{ref:'A2',raw:'100'},{ref:'B2',raw:'5'},{ref:'A3',raw:'100'},{ref:'B3',raw:'7'},{ref:'A4',raw:'101'},{ref:'B4',raw:'3'}],false);
 const model=buildDataModel([{name:'Dates',book:dates},{name:'Calendar',book:calendar},{name:'Sales',book:sales}],{
  tables:[{name:'Dates',sheet:'Dates',range:'A1:B3'},{name:'Calendar',sheet:'Calendar',range:'A1:B3'},{name:'Sales',sheet:'Sales',range:'A1:B4'}],
  relationships:[{manyTable:'Calendar',manyColumn:'DateID',oneTable:'Dates',oneColumn:'DateID'},{manyTable:'Sales',manyColumn:'CalendarID',oneTable:'Calendar',oneColumn:'CalendarID'}],measures:[],
 });
 assert.equal(evaluateDaxMeasure(model,'SUM(Sales[Amount])',[{table:'Dates',column:'DateID',values:[1]}]),12);
 assert.equal(evaluateDaxMeasure(model,'SUMX(Sales, Sales[Amount] * RELATED(Dates[Year]))'),30363);
});

test('local DAX CALCULATE supports KEEPFILTERS and bounded REMOVEFILTERS/ALL modifiers',()=>{
 const model=runtime(),east={table:'Sales',column:'Region',values:['East']};
 assert.equal(evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), REMOVEFILTERS(Sales[Region]))',[east]),12);
 assert.equal(evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), ALL(Sales))',[east]),12);
 assert.equal(evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), REMOVEFILTERS())',[east]),12);
 assert.equal(evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), KEEPFILTERS(Sales[Region] = "West"))',[east]),null);
 assert.equal(evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), KEEPFILTERS(Sales[Region] = "West"))'),5);
 assert.throws(()=>evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), ALL())'),/needs a model table or column/);
});

test('local DAX ALLEXCEPT keeps selected direct filters and clears other table filters',()=>{
 const model=runtime(),filters=[{table:'Sales',column:'Region',values:['East']},{table:'Sales',column:'Amount',values:[0]}];
 assert.equal(evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), ALLEXCEPT(Sales, Sales[Region]))',filters),7);
 assert.throws(()=>evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), ALLEXCEPT(Sales, Other[Region]))'),/belong to its first table/);
 assert.throws(()=>evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), ALLEXCEPT(Sales))'),/expects a model table/);
});

test('local DAX iterators support Boolean NOT, AND and OR with precedence',()=>{
 const model=runtime();
 assert.equal(evaluateDaxMeasure(model,'SUMX(Sales, IF(Sales[Amount] > 5 && Sales[Region] = "East", 1, 0))'),1);
 assert.equal(evaluateDaxMeasure(model,'SUMX(Sales, IF(NOT (Sales[Amount] > 5) || Sales[Region] = "West", 1, 0))'),3);
 assert.equal(evaluateDaxMeasure(model,'IF(TRUE || FALSE && FALSE, 1, 0)'),1);
 assert.equal(evaluateDaxMeasure(model,'IF(AND(TRUE, NOT(FALSE)), 8, 0)'),8);
});

test('local DAX adds bounded numeric rounding, remainder, power and root functions',()=>{
 const model=runtime();
 assert.equal(evaluateDaxMeasure(model,'INT(-1.2)'),-2);
 assert.equal(evaluateDaxMeasure(model,'TRUNC(-1.29, 1)'),-1.2);
 assert.equal(evaluateDaxMeasure(model,'ROUNDUP(-1.231, 2)'),-1.24);
 assert.equal(evaluateDaxMeasure(model,'ROUNDDOWN(-1.239, 2)'),-1.23);
 assert.equal(evaluateDaxMeasure(model,'MOD(-5, 3)'),1);
 assert.equal(evaluateDaxMeasure(model,'MOD(5, -3)'),-1);
 assert.equal(evaluateDaxMeasure(model,'POWER(3, 2) + SQRT(16)'),13);
 assert.equal(evaluateDaxMeasure(model,'SUMX(Sales, IF(Sales[Region] = "West", SQRT(Sales[Amount]), 0))'),Math.sqrt(5));
 assert.throws(()=>evaluateDaxMeasure(model,'MOD(1, 0)'),/cannot be zero/);
 assert.throws(()=>evaluateDaxMeasure(model,'SQRT(-1)'),/non-negative/);
});

test('local DAX supports bounded text helpers in measures and iterators',()=>{
 const model=runtime();
 assert.equal(evaluateDaxMeasure(model,'LEN(BLANK())'),0);
 assert.equal(evaluateDaxMeasure(model,'LEFT("GridForge", 4) & RIGHT("Sheet", 2)'), 'Gridet');
 assert.equal(evaluateDaxMeasure(model,'TRIM("  grid   forge  ")'),'grid forge');
 assert.equal(evaluateDaxMeasure(model,'UPPER("west")'),'WEST');
 assert.equal(evaluateDaxMeasure(model,'CONCATENATE("grid", "forge")'),'gridforge');
 assert.equal(evaluateDaxMeasure(model,'SUMX(Sales, IF(UPPER(Sales[Region]) = "WEST", 1, 0))'),1);
 assert.throws(()=>evaluateDaxMeasure(model,'LEFT("text", -1)'),/character count/);
});

test('local DAX CONCATENATEX joins a bounded row expression in source order',()=>{
 const model=runtime();
 assert.equal(evaluateDaxMeasure(model,'CONCATENATEX(Sales, Sales[Region])'),'EastWestEast');
 assert.equal(evaluateDaxMeasure(model,'CONCATENATEX(Sales, Sales[Region], " | ")'),'East | West | East | ');
 assert.equal(evaluateDaxMeasure(model,'CONCATENATEX(FILTER(Sales, Sales[Region] = "West"), UPPER(Sales[Region]), ";")'),'WEST');
 assert.equal(evaluateDaxMeasure(model,'CONCATENATEX(FILTER(Sales, Sales[Region] = "missing"), Sales[Region], ",")'),null);
 assert.throws(()=>evaluateDaxMeasure(model,'CONCATENATEX(Sales, Sales[Region], ",", Sales[Amount])'),/order-by arguments are not supported/);
 assert.throws(()=>evaluateDaxMeasure(model,'CONCATENATEX(Sales, Sales[Missing])'),/Unknown model column/);
 const large=runtime(),largeTable=large.tables.get('sales');assert.ok(largeTable);largeTable.columns=['Label'];largeTable.rows=[['x'.repeat(16_384)],['y'.repeat(16_384)]];
 assert.throws(()=>evaluateDaxMeasure(large,'CONCATENATEX(Sales, Sales[Label], "|")'),/32,767-character cell limit/);
});

test('local DAX RANKX supports bounded numeric ranks, order and tie modes',()=>{
 const model=runtime();
 assert.equal(evaluateDaxMeasure(model,'RANKX(Sales, Sales[Amount], 5)'),2);
 assert.equal(evaluateDaxMeasure(model,'RANKX(Sales, Sales[Amount], 5, ASC, DENSE)'),2);
 assert.equal(evaluateDaxMeasure(model,'RANKX(Sales, Sales[Amount], -1, DESC, SKIP)'),5);
 assert.equal(evaluateDaxMeasure(model,'RANKX(Sales, Sales[Amount], -1, DESC, DENSE)'),4);
 assert.equal(evaluateDaxMeasure(model,'RANKX(FILTER(Sales, Sales[Amount] > 0), Sales[Amount], 7, DESC, DENSE)'),1);
 assert.equal(evaluateDaxMeasure(model,'RANKX(FILTER(Sales, Sales[Amount] > 100), Sales[Amount], 5)'),null);
 assert.throws(()=>evaluateDaxMeasure(model,'RANKX(Sales, Sales[Region], 1)'),/must return numbers or BLANK/);
 assert.throws(()=>evaluateDaxMeasure(model,'RANKX(Sales, Sales[Amount], 1, SIDEWAYS)'),/order must be ASC\/1 or DESC\/0/);
});

test('local DAX percentile inclusive functions interpolate numeric model values',()=>{
 const model=runtime();
 assert.equal(evaluateDaxMeasure(model,'PERCENTILE.INC(Sales[Amount], 0.5)'),5);
 assert.equal(evaluateDaxMeasure(model,'PERCENTILE.INC(Sales[Amount], 0.25)'),2.5);
 assert.equal(evaluateDaxMeasure(model,'PERCENTILEX.INC(Sales, Sales[Amount], 0.5)'),5);
 assert.equal(evaluateDaxMeasure(model,'PERCENTILEX.INC(FILTER(Sales, Sales[Amount] > 0), Sales[Amount], 0.25)'),5.5);
 assert.equal(evaluateDaxMeasure(model,'PERCENTILE.INC(Sales[Amount], 0.5)',[{table:'Sales',column:'Region',values:['missing']}]),null);
 assert.throws(()=>evaluateDaxMeasure(model,'PERCENTILE.INC(Sales[Amount], 1.1)'),/percentile must be a number from 0 through 1/);
 assert.equal(evaluateDaxMeasure(model,'PERCENTILEX.INC(Sales, Sales[Region], 0.5)'),null);
});

test('local DAX sample and population variance and standard deviation aggregate columns and iterators',()=>{
 const model=runtime(),close=(actual,expected)=>assert.ok(Math.abs(actual-expected)<1e-12,`${actual} != ${expected}`);
 close(evaluateDaxMeasure(model,'VAR.P(Sales[Amount])'),26/3);
 close(evaluateDaxMeasure(model,'VAR.S(Sales[Amount])'),13);
 close(evaluateDaxMeasure(model,'STDEV.P(Sales[Amount])'),Math.sqrt(26/3));
 close(evaluateDaxMeasure(model,'STDEV.S(Sales[Amount])'),Math.sqrt(13));
 close(evaluateDaxMeasure(model,'VARX.P(Sales, Sales[Amount])'),26/3);
 close(evaluateDaxMeasure(model,'VARX.S(Sales, Sales[Amount])'),13);
 close(evaluateDaxMeasure(model,'STDEVX.P(Sales, Sales[Amount])'),Math.sqrt(26/3));
 close(evaluateDaxMeasure(model,'STDEVX.S(Sales, Sales[Amount])'),Math.sqrt(13));
 assert.equal(evaluateDaxMeasure(model,'VAR.P(Sales[Amount])',[{table:'Sales',column:'Amount',values:[0]}]),0);
 assert.equal(evaluateDaxMeasure(model,'STDEV.P(Sales[Amount])',[{table:'Sales',column:'Amount',values:['No match']}]),null);
 assert.throws(()=>evaluateDaxMeasure(model,'VAR.S(Sales[Amount])',[{table:'Sales',column:'Amount',values:[0]}]),/at least two nonblank numeric values/);
 assert.throws(()=>evaluateDaxMeasure(model,'STDEVX.S(Sales, Sales[Amount])',[{table:'Sales',column:'Region',values:['West']}]),/at least two nonblank numeric values/);
});
test('local DAX percentile exclusive functions enforce exclusive rank bounds',()=>{
 const model=runtime();
 assert.equal(evaluateDaxMeasure(model,'PERCENTILE.EXC(Sales[Amount], 0.5)'),5);
 assert.ok(Math.abs(evaluateDaxMeasure(model,'PERCENTILE.EXC(Sales[Amount], 0.4)')-3)<1e-12);
 assert.equal(evaluateDaxMeasure(model,'PERCENTILEX.EXC(Sales, Sales[Amount], 0.5)'),5);
 assert.equal(evaluateDaxMeasure(model,'PERCENTILEX.EXC(FILTER(Sales, Sales[Amount] > 0), Sales[Amount], 0.5)'),6);
 assert.throws(()=>evaluateDaxMeasure(model,'PERCENTILE.EXC(Sales[Amount], 0)'),/greater than 0 and less than 1/);
 assert.throws(()=>evaluateDaxMeasure(model,'PERCENTILE.EXC(Sales[Amount], 0.1)'),/outside the range/);
});

test('local DAX exposes bounded Gregorian date parts for ISO dates and Excel serials',()=>{
 const model=runtime();
 assert.equal(evaluateDaxMeasure(model,'YEAR("2024-02-29")'),2024);
 assert.equal(evaluateDaxMeasure(model,'MONTH("2024-02-29")'),2);
 assert.equal(evaluateDaxMeasure(model,'DAY("2024-02-29")'),29);
 assert.equal(evaluateDaxMeasure(model,'QUARTER("2024-02-29")'),1);
 assert.equal(evaluateDaxMeasure(model,'WEEKDAY("2024-02-29")'),5);
 assert.equal(evaluateDaxMeasure(model,'WEEKDAY("2024-02-29", 2)'),4);
 assert.equal(evaluateDaxMeasure(model,'WEEKDAY("2024-02-29", 3)'),3);
 assert.equal(evaluateDaxMeasure(model,'WEEKNUM("2024-02-29")'),9);
 assert.equal(evaluateDaxMeasure(model,'WEEKNUM("2024-02-29", 11)'),9);
 assert.equal(evaluateDaxMeasure(model,'WEEKNUM("2024-02-29", 21)'),9);
 assert.equal(evaluateDaxMeasure(model,'WEEKNUM("2021-01-01", 21)'),53);
 assert.equal(evaluateDaxMeasure(model,'WEEKNUM("2010-02-14", 2)'),7);
 assert.equal(evaluateDaxMeasure(model,'WEEKNUM("2021-01-01", 1)'),1);
 assert.equal(evaluateDaxMeasure(model,'YEAR(60)'),1900);
 assert.equal(evaluateDaxMeasure(model,'MONTH(60)'),2);
 assert.equal(evaluateDaxMeasure(model,'DAY(60)'),29);
 assert.equal(evaluateDaxMeasure(model,'DATE(2024, 13, 1)'),excelDateSerial('2025-01-01'));
 assert.equal(evaluateDaxMeasure(model,'DATE(1900, 2, 29)'),60);
 assert.equal(evaluateDaxMeasure(model,'YEAR(DATE(1900, 2, 29))'),1900);
 assert.equal(evaluateDaxMeasure(model,'DATE(1899, 12, 31)'),excelDateSerial('3799-12-31'));
 assert.equal(evaluateDaxMeasure(model,'WEEKDAY(60, 2)'),4);
 assert.equal(evaluateDaxMeasure(model,'YEAR(BLANK())'),null);
 assert.throws(()=>evaluateDaxMeasure(model,'MONTH("03/04/2024")'),/unambiguous ISO date/);
 assert.throws(()=>evaluateDaxMeasure(model,'WEEKDAY("2024-02-29", 4)'),/return type must be 1, 2, or 3/);
 assert.throws(()=>evaluateDaxMeasure(model,'WEEKNUM("2024-02-29", 22)'),/return type must be 1, 2, or 11 through 17, or 21/);
 assert.equal(evaluateDaxMeasure(model,'SUMX(Sales, IF(YEAR(Sales[Amount]) = 1900, 1, 0))'),2);
 assert.equal(evaluateDaxMeasure(model,'SUMX(Sales, WEEKDAY("2024-02-29", 2))'),16);
 assert.equal(evaluateDaxMeasure(model,'SUMX(Sales, WEEKNUM("2024-02-29", 21))'),36);
 assert.equal(evaluateDaxMeasure(model,'SUMX(Sales, YEAR(DATE(2024, 1, 1)))'),8096);
 assert.throws(()=>evaluateDaxMeasure(model,'DATE(-1, 1, 1)'),/calendar range/);
});

test('local DAX YTD/QTD/MTD totals expand date filters from the latest visible ISO date',()=>{
 const calendar=new Workbook(12,4),sales=new Workbook(12,3);
 calendar.applyEdits([{ref:'A1',raw:'Date'},{ref:'B1',raw:'Year'},{ref:'C1',raw:'Month'},{ref:'A2',raw:'2024-01-01'},{ref:'B2',raw:'2024'},{ref:'C2',raw:'1'},{ref:'A3',raw:'2024-01-15'},{ref:'B3',raw:'2024'},{ref:'C3',raw:'1'},{ref:'A4',raw:'2024-02-01'},{ref:'B4',raw:'2024'},{ref:'C4',raw:'2'},{ref:'A5',raw:'2024-03-15'},{ref:'B5',raw:'2024'},{ref:'C5',raw:'3'},{ref:'A6',raw:'2024-04-01'},{ref:'B6',raw:'2024'},{ref:'C6',raw:'4'},{ref:'A7',raw:'2025-01-01'},{ref:'B7',raw:'2025'},{ref:'C7',raw:'1'}],false);
 sales.applyEdits([{ref:'A1',raw:'Date'},{ref:'B1',raw:'Amount'},{ref:'A2',raw:'2024-01-01'},{ref:'B2',raw:'5'},{ref:'A3',raw:'2024-01-15'},{ref:'B3',raw:'7'},{ref:'A4',raw:'2024-02-01'},{ref:'B4',raw:'4'},{ref:'A5',raw:'2024-03-15'},{ref:'B5',raw:'6'},{ref:'A6',raw:'2024-04-01'},{ref:'B6',raw:'11'},{ref:'A7',raw:'2025-01-01'},{ref:'B7',raw:'13'}],false);
 const model=buildDataModel([{name:'Dates',book:calendar},{name:'Sales',book:sales}],{tables:[{name:'Dates',sheet:'Dates',range:'A1:C7'},{name:'Sales',sheet:'Sales',range:'A1:B7'}],relationships:[{manyTable:'Sales',manyColumn:'Date',oneTable:'Dates',oneColumn:'Date'}],measures:[]});
 assert.equal(evaluateDaxMeasure(model,'TOTALYTD(SUM(Sales[Amount]), Dates[Date])',[{table:'Dates',column:'Date',values:['2024-01-15']}]),12);
 assert.equal(evaluateDaxMeasure(model,'TOTALYTD(SUM(Sales[Amount]), Dates[Date])',[{table:'Dates',column:'Month',values:[2]}]),16);
 assert.equal(evaluateDaxMeasure(model,'TOTALQTD(SUM(Sales[Amount]), Dates[Date])',[{table:'Dates',column:'Date',values:['2024-04-01']}]),11);
 assert.equal(evaluateDaxMeasure(model,'TOTALMTD(SUM(Sales[Amount]), Dates[Date])',[{table:'Dates',column:'Date',values:['2024-04-01']}]),11);
 assert.equal(evaluateDaxMeasure(model,'TOTALYTD(SUM(Sales[Amount]), Dates[Date])',[{table:'Dates',column:'Year',values:[2025]}]),13);
 assert.equal(evaluateDaxMeasure(model,'TOTALQTD(SUM(Sales[Amount]), Dates[Date])',[{table:'Dates',column:'Month',values:[3]}]),22);
 assert.equal(evaluateDaxMeasure(model,'TOTALMTD(SUM(Sales[Amount]), Dates[Date])',[{table:'Dates',column:'Month',values:[3]}]),6);
 assert.equal(evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), DATESYTD(Dates[Date]))',[{table:'Dates',column:'Date',values:['2024-01-15']}]),12);
 assert.equal(evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), DATESQTD(Dates[Date]))',[{table:'Dates',column:'Month',values:[3]}]),22);
 assert.equal(evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), DATESMTD(Dates[Date]))',[{table:'Dates',column:'Month',values:[3]}]),6);
 assert.equal(evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), DATESBETWEEN(Dates[Date], "2024-01-15", "2024-03-15"))',[{table:'Dates',column:'Month',values:[4]}]),17);
 assert.equal(evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), DATESINPERIOD(Dates[Date], "2024-03-15", -2, MONTH))',[{table:'Dates',column:'Month',values:[4]}]),17);
 assert.equal(evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), DATESINPERIOD(Dates[Date], "2024-03-15", 1, MONTH))'),17);
 assert.equal(evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), DATESINPERIOD(Dates[Date], "2024-03-15", -7, DAY))'),6);
 assert.equal(evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), DATESINPERIOD(Dates[Date], "2024-03-15", 0, DAY))'),6);
 assert.equal(evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), DATEADD(Dates[Date], -3, MONTH))',[{table:'Dates',column:'Date',values:['2024-04-01']}]),5);
 assert.equal(evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), DATEADD(Dates[Date], 0, DAY))',[{table:'Dates',column:'Month',values:[4]}]),11);
 assert.equal(evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), SAMEPERIODLASTYEAR(Dates[Date]))',[{table:'Dates',column:'Date',values:['2025-01-01']}]),5);
 assert.equal(evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), PARALLELPERIOD(Dates[Date], -1, MONTH))',[{table:'Dates',column:'Date',values:['2024-04-01']}]),6);
 assert.equal(evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), PARALLELPERIOD(Dates[Date], -1, QUARTER))',[{table:'Dates',column:'Date',values:['2024-04-01']}]),22);
 assert.equal(evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), PARALLELPERIOD(Dates[Date], -1, YEAR))',[{table:'Dates',column:'Date',values:['2025-01-01']}]),33);
 assert.equal(evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), DATESBETWEEN(Dates[Date], BLANK(), BLANK()))'),46);
 assert.equal(evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), DATESBETWEEN(Dates[Date], "2024-03-15", "2024-01-15"))'),null);
 assert.equal(evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), Dates[Year] = 2025, DATESBETWEEN(Dates[Date], "2024-01-01", "2024-03-15"))'),null);
 assert.equal(evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), Dates[Year] = 2025, DATESQTD(Dates[Date]))',[{table:'Dates',column:'Month',values:[3]}]),null);
 assert.equal(evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), DATESQTD(Dates[Date]), Dates[Year] = 2025)',[{table:'Dates',column:'Month',values:[3]}]),null);
 assert.equal(evaluateDaxMeasure(model,'TOTALYTD(SUM(Sales[Amount]), Dates[Date])',[{table:'Dates',column:'Date',values:['missing']}]),null);
 assert.equal(evaluateDaxMeasure(model,'TOTALYTD(SUM(Sales[Amount]), Sales[Date])'),13);
 assert.throws(()=>evaluateDaxMeasure(model,'TOTALYTD(SUM(Sales[Amount]))'),/one date-column reference/);
 assert.throws(()=>evaluateDaxMeasure(model,'TOTALQTD(SUM(Sales[Amount]), Dates[Date], 1)'),/one date-column reference/);
 assert.throws(()=>evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), DATESYTD("2024-01-01"))'),/one date-column reference/);
 assert.throws(()=>evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), DATESBETWEEN(Dates[Date], "03/04/2024", "2024-03-05"))'),/bounds must be BLANK/);
 assert.throws(()=>evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), DATESINPERIOD(Dates[Date], "2024-03-15", 1, WEEK))'),/interval must be DAY/);
 assert.throws(()=>evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), DATESINPERIOD(Dates[Date], "2024-03-15", 1.5, DAY))'),/interval count must be an integer/);
 assert.throws(()=>evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), DATEADD(Dates[Date], 0.5, DAY))'),/interval count must be an integer/);
 assert.throws(()=>evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), SAMEPERIODLASTYEAR(Dates[Date], 1))'),/one date-column reference/);
 assert.throws(()=>evaluateDaxMeasure(model,'CALCULATE(SUM(Sales[Amount]), PARALLELPERIOD(Dates[Date], -1, DAY))'),/interval must be MONTH/);
 const serialDates=new Workbook(5,2),serialSales=new Workbook(5,2),jan1=excelDateSerial('2024-01-01'),jan2=excelDateSerial('2024-01-02');
 serialDates.applyEdits([{ref:'A1',raw:'Date'},{ref:'A2',raw:String(jan1)},{ref:'A3',raw:String(jan2)}],false);serialSales.applyEdits([{ref:'A1',raw:'Date'},{ref:'B1',raw:'Amount'},{ref:'A2',raw:String(jan1)},{ref:'B2',raw:'2'},{ref:'A3',raw:String(jan2)},{ref:'B3',raw:'5'}],false);
 const serialModel=buildDataModel([{name:'Calendar',book:serialDates},{name:'Facts',book:serialSales}],{tables:[{name:'Calendar',sheet:'Calendar',range:'A1:A3'},{name:'Facts',sheet:'Facts',range:'A1:B3'}],relationships:[{manyTable:'Facts',manyColumn:'Date',oneTable:'Calendar',oneColumn:'Date'}],measures:[]});
 assert.equal(evaluateDaxMeasure(serialModel,'TOTALYTD(SUM(Facts[Amount]), Calendar[Date])',[{table:'Calendar',column:'Date',values:[jan2]}]),7);
 const timedDates=new Workbook(5,2),timedFacts=new Workbook(5,2),feb=excelDateSerial('2024-02-01')+0.5,mar1=excelDateSerial('2024-03-01')+0.75,mar31=excelDateSerial('2024-03-31')+0.5;timedDates.applyEdits([{ref:'A1',raw:'Date'},{ref:'A2',raw:String(feb)},{ref:'A3',raw:String(mar1)},{ref:'A4',raw:String(mar31)}],false);timedFacts.applyEdits([{ref:'A1',raw:'Date'},{ref:'B1',raw:'Amount'},{ref:'A2',raw:String(feb)},{ref:'B2',raw:'2'},{ref:'A3',raw:String(mar1)},{ref:'B3',raw:'3'},{ref:'A4',raw:String(mar31)},{ref:'B4',raw:'5'}],false);const timedModel=buildDataModel([{name:'Calendar',book:timedDates},{name:'Facts',book:timedFacts}],{tables:[{name:'Calendar',sheet:'Calendar',range:'A1:A4'},{name:'Facts',sheet:'Facts',range:'A1:B4'}],relationships:[{manyTable:'Facts',manyColumn:'Date',oneTable:'Calendar',oneColumn:'Date'}],measures:[]});
 assert.equal(evaluateDaxMeasure(timedModel,'CALCULATE(SUM(Facts[Amount]), PARALLELPERIOD(Calendar[Date], -1, MONTH))',[{table:'Calendar',column:'Date',values:[mar1,mar31]}]),2);
});
