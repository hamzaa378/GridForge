import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook} from '../dist/engine/workbook.js';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {buildQueryRefresh,validateQuerySpec} from '../dist/engine/power_query.js';
import {executePowerQueryM} from '../dist/engine/power_query_m.js';
import {excelDateSerial} from '../dist/engine/date_serial.js';
import {validateTools} from '../dist/engine/sheettools.js';
import {exportXlsx} from '../dist/engine/xlsx.js';
import {zipRead} from '../dist/engine/zip.js';

function sample(){const book=new Workbook(30,12);book.applyEdits([{ref:'A1',raw:' Name '},{ref:'B1',raw:'Amount'},{ref:'A2',raw:' alice '},{ref:'B2',raw:'12'},{ref:'A3',raw:'Bob'},{ref:'B3',raw:'7'}],false);return book;}
test('bounded M column selection supports explicit missing-field behavior and promotes headers',()=>{
 const result=executePowerQueryM([['Input','Value'],['Name','Amount'],['Ada',3],['Lin',4]],'let Selected = Table.SelectColumns(Source, {"Input", "Missing"}, MissingField.UseNull), Promoted = Table.PromoteHeaders(Selected) in Promoted');
 assert.deepEqual(result,[['Name','Column2'],['Ada',null],['Lin',null]]);
 assert.throws(()=>executePowerQueryM([['A'],[1]],'let Selected = Table.SelectColumns(Source, {"Missing"}) in Selected'),/missing column/);
});
test('bounded M reorders columns, demotes headers, and reports table metadata',()=>{
 const result=executePowerQueryM([['Name','Amount'],['Ada',3]],'let Reordered = Table.ReorderColumns(Source, {"Amount", "Unknown", "Name"}, MissingField.Ignore), Demoted = Table.DemoteHeaders(Reordered), Names = Table.ColumnNames(Demoted), Count = Table.RowCount(Demoted) in Demoted');
 assert.deepEqual(result,[['Column1','Column2'],['Amount','Name'],[3,'Ada']]);
 assert.deepEqual(executePowerQueryM([['Name','Amount'],['Ada',3]],'let Names = Table.ColumnNames(Source) in #table({"Column"}, {{List.First(Names)}})'),[['Column'],['Name']]);
 assert.deepEqual(executePowerQueryM([['Name'],['Ada'],['Lin']],'let Count = Table.RowCount(Source) in #table({"Rows"}, {{Count}})'),[['Rows'],[2]]);
});
test('bounded M exposes table column lists, column-presence checks, and empty-table state',()=>{
 const input=[['Team','Amount'],['A',4],['B',3]];
 assert.deepEqual(executePowerQueryM(input,'let Amounts = Table.Column(Source, "Amount"), Result = #table({"Total", "HasBoth", "HasMissing", "Empty"}, {{List.Sum(Amounts), Table.HasColumns(Source, {"Team", "Amount"}), Table.HasColumns(Source, "Missing"), Table.IsEmpty(Source)}}) in Result'),[['Total','HasBoth','HasMissing','Empty'],[7,true,false,false]]);
 assert.deepEqual(executePowerQueryM([['Amount']],'let Values = Table.Column(Source, "Amount"), Result = #table({"Count", "Empty", "HasNoColumns"}, {{List.Count(Values), Table.IsEmpty(Source), Table.HasColumns(Source, {})}}) in Result'),[['Count','Empty','HasNoColumns'],[0,true,true]]);
 assert.throws(()=>executePowerQueryM(input,'let Values = Table.Column(Source, "Missing") in #table({"Count"}, {{List.Count(Values)}})'),/cannot find column/);
});
test('bounded M reports table column counts and materialized row-count estimates',()=>{
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let T = #table({"A","B"},{{1,2},{3,4}}), Empty = #table({"Only"},{}), Result = #table({"Rows","Estimate","Columns","EmptyRows","EmptyColumns"}, {{Table.RowCount(T),Table.ApproximateRowCount(T),Table.ColumnCount(T),Table.ApproximateRowCount(Empty),Table.ColumnCount(Empty)}}) in Result'),[['Rows','Estimate','Columns','EmptyRows','EmptyColumns'],[2,2,2,0,1]]);
 assert.deepEqual(executePowerQueryM([['2','1'],['two','one']],'let Ok = Table.MatchesAllRows(Source, each List.First(Record.FieldNames(_)) = "2") in #table({"Ordered"}, {{Ok}})'),[['Ordered'],[true]]);
});
test('bounded M table predicate summaries short-circuit and use empty-table identities',()=>{
 const input=[['N'],[1],[2],[3]];
 assert.deepEqual(executePowerQueryM(input,'let All = Table.MatchesAllRows(Source, each [N] > 0), Any = Table.MatchesAnyRows(Source, each [N] = 2), None = Table.MatchesAnyRows(Source, each [N] < 0) in #table({"All","Any","None"}, {{All,Any,None}})'),[['All','Any','None'],[true,true,false]]);
 assert.deepEqual(executePowerQueryM([['N']],'let All = Table.MatchesAllRows(Source, each [N] > 0), Any = Table.MatchesAnyRows(Source, each [N] > 0) in #table({"All","Any"}, {{All,Any}})'),[['All','Any'],[true,false]]);
 assert.throws(()=>executePowerQueryM(input,'let Value = Table.MatchesAllRows(Source, each [N]) in #table({"Value"}, {{Value}})'),/predicates must return true or false/);
});
test('bounded M Table.Profile summarizes typed scalar columns and null values',()=>{
 assert.deepEqual(executePowerQueryM([['Name','Value'],['a',1],['a',null],['b',1]],'let Profile = Table.Profile(Source) in Profile'),[['Column','Min','Max','Count','NullCount','DistinctCount'],['Name','a','b',3,0,2],['Value',1,1,2,1,2]]);
 assert.deepEqual(executePowerQueryM([['Empty']],'let Profile = Table.Profile(Source) in Profile'),[['Column','Min','Max','Count','NullCount','DistinctCount'],['Empty',null,null,0,0,0]]);
});
test('bounded M Table.First, Table.Last and Table.FirstValue support bounded defaults',()=>{
 const result=executePowerQueryM([['Seed'],[0]],'let Local = #table({"A","B"}, {{1,"x"},{2,"y"}}), First = Table.First(Local), Last = Table.Last(Local), Rows = Table.FromRecords({First,Last}) in Rows');
 assert.deepEqual(result,[['A','B'],[1,'x'],[2,'y']]);
 assert.equal(executePowerQueryM([['Seed'],[0]],'let Value = Table.FirstValue(#table({"A"}, {{5}})) in #table({"Value"}, {{Value}})')[1][0],5);
 const empty=executePowerQueryM([['Seed'],[0]],'let Local = #table({"A"}, {}), First = Table.First(Local, [A=7]), Last = Table.Last(Local, [A=9]), Rows = Table.FromRecords({First,Last}) in Rows');
 assert.deepEqual(empty,[['A'],[7],[9]]);
 assert.equal(executePowerQueryM([['Seed'],[0]],'let Value = Table.FirstValue(#table({"A"}, {}), "empty") in #table({"Value"}, {{Value}})')[1][0],'empty');
 assert.throws(()=>executePowerQueryM([['Seed'],[0]],'let Value = Table.First(#table({"A"}, {})) in #table({"Value"}, {{Value}})'),/empty table without a default/);
});
test('bounded M List.Product and List.Median work in scalar and grouping aggregations',()=>{
 const input=[['Team','Kind','Amount'],['A','x',2],['A','x',3],['A','y',4],['B','x',5]];
 assert.deepEqual(executePowerQueryM(input,'let Values = Table.Column(Source, "Amount"), Stats = #table({"Product", "Median", "EmptyProduct"}, {{List.Product(Values), List.Median(Values), List.Product({})}}) in Stats'),[['Product','Median','EmptyProduct'],[120,3.5,null]]);
 assert.deepEqual(executePowerQueryM(input,'let Grouped = Table.Group(Source, {"Team"}, {{"Product", each List.Product([Amount])}, {"Median", each List.Median([Amount])}}) in Grouped'),[['Team','Product','Median'],['A',24,3],['B',5,5]]);
 assert.deepEqual(executePowerQueryM(input,'let Pivot = Table.Pivot(Source, {"x", "y"}, "Kind", "Amount", List.Product) in Pivot'),[['Team','x','y'],['A',6,4],['B',5,null]]);
 assert.deepEqual(executePowerQueryM([['Value'],[null],[null]],'let Values = Table.Column(Source, "Value"), Stats = #table({"Product", "Median"}, {{List.Product(Values), List.Median(Values)}}) in Stats'),[['Product','Median'],[null,null]]);
 assert.throws(()=>executePowerQueryM([['Value'],['text']],'let Values = Table.Column(Source, "Value"), Product = List.Product(Values) in #table({"Value"}, {{Product}})'),/requires numeric list values/);
 assert.throws(()=>executePowerQueryM([['Team','Kind','Value'],['A','x',1e308],['A','x',1e308]],'let Pivot = Table.Pivot(Source, {"x"}, "Kind", "Value", List.Product) in Pivot'),/non-finite result/);
});
test('bounded M List.StandardDeviation computes sample deviation in scalar and pivot aggregation',()=>{
 assert.ok(Math.abs(executePowerQueryM([['Seed'],[1]],'let Result = #table({"Std"}, {{List.StandardDeviation({1, 2, 3, 4, 5})}}) in Result')[1][0]-Math.sqrt(2.5))<1e-12);
 assert.deepEqual(executePowerQueryM([['Team','Kind','Value'],['A','x',1],['A','x',3],['A','y',2],['A','y',6]],'let Pivot = Table.Pivot(Source, {"x", "y"}, "Kind", "Value", List.StandardDeviation) in Pivot'),[['Team','x','y'],['A',Math.sqrt(2),Math.sqrt(8)]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Value = List.StandardDeviation({1}) in #table({"Value"}, {{Value}})'),/at least two numeric values/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Value = List.StandardDeviation({1, null}) in #table({"Value"}, {{Value}})'),/every list item to be a finite number/);
});
test('bounded M List.Mode and List.Modes preserve type-aware ties and source ordering',()=>{
 const query='let Modes = List.Modes({"A", 1, 1, "A", true}), Tied = List.Modes({2, 1, 2, 1}), Selected = List.Mode({2, 1, 2, 1}), Result = #table({"Mode", "ModesFirst", "ModesLast", "TiedCount", "TiedFirst", "Selected"}, {{List.Mode({"A", 1, 1, "A", true}), List.First(Modes), List.Last(Modes), List.Count(Tied), List.First(Tied), Selected}}) in Result';
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],query),[['Mode','ModesFirst','ModesLast','TiedCount','TiedFirst','Selected'],['A','A',1,2,2,1]]);
 assert.deepEqual(executePowerQueryM([['Team','Kind','Value'],['A','x',1],['A','x',2],['A','y',3],['A','y',4]],'let Pivot = Table.Pivot(Source, {"x", "y"}, "Kind", "Value", List.Mode) in Pivot'),[['Team','x','y'],['A',2,4]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Value = List.Mode({}) in #table({"Value"}, {{Value}})'),/empty list/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Value = List.Modes({1}, each _) in #table({"Value"}, {{Value}})'),/expects 1 argument/);
});
test('bounded M List.Percentile supports Excel inclusive and exclusive interpolation',()=>{
 const expr='let Inclusive = List.Percentile({5, 3, 1, 7, 9}, 0.25), InclusiveMany = List.Percentile({5, 3, 1, 7, 9}, {0.25, 0.5, 0.75}), Exclusive = List.Percentile({5, 3, 1, 7, 9}, {0.25, 0.5, 0.75}, [PercentileMode = PercentileMode.ExcelExc]), ExclusiveMiddle = List.Percentile({5, 3, 1, 7, 9}, 0.5, [PercentileMode = PercentileMode.ExcelExc]), Result = #table({"Inc", "IncMid", "ExcFirst", "ExcMid", "ExcLast"}, {{Inclusive, List.First(InclusiveMany), List.First(Exclusive), ExclusiveMiddle, List.Last(Exclusive)}}) in Result';
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],expr),[['Inc','IncMid','ExcFirst','ExcMid','ExcLast'],[3,3,2,5,8]]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let P = List.Percentile({1, 2, 3, 4}, 0.5, [PercentileMode = PercentileMode.ExcelExc]), Result = #table({"Value"}, {{P}}) in Result'),[['Value'],[2.5]]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let Values = List.Percentile({}, {0.25, 0.75}), Result = #table({"Count"}, {{List.Count(Values)}}) in Result'),[['Count'],[2]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let P = List.Percentile({1, 2, 3}, 0.1, [PercentileMode = PercentileMode.ExcelExc]), Result = #table({"Value"}, {{P}}) in Result'),/falls outside the available numeric values/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let P = List.Percentile({1, 2, 3}, 1.1), Result = #table({"Value"}, {{P}}) in Result'),/probability must be from 0 through 1/);
});
test('bounded M List.Sort handles explicit order and rejects mixed incomparable values',()=>{
 const query='let Asc = List.Sort({3, 1, 2}), Desc = List.Sort({3, 1, 2}, Order.Descending), Result = #table({"AscFirst", "AscLast", "DescFirst", "DescLast", "NullFirst"}, {{List.First(Asc), List.Last(Asc), List.First(Desc), List.Last(Desc), List.First(List.Sort({3, null, 1}))}}) in Result';
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],query),[['AscFirst','AscLast','DescFirst','DescLast','NullFirst'],[1,3,3,1,null]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Values = List.Sort({1, "a"}), Result = #table({"Count"}, {{List.Count(Values)}}) in Result'),/one scalar type/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Values = List.Sort({1, 2}, 0), Result = #table({"Count"}, {{List.Count(Values)}}) in Result'),/Order.Ascending or Order.Descending/);
});
test('bounded M list slicing handles first, last, skip and ranged windows',()=>{
 const query='let First = List.FirstN({1, 2, 3, 4}, 2), Last = List.LastN({1, 2, 3, 4}, 2), Skipped = List.Skip({1, 2, 3, 4}, 2), Ranged = List.Range({1, 2, 3, 4}, 1, 2), NoFirst = List.RemoveFirstN({1, 2, 3, 4}, 2), NoLast = List.RemoveLastN({1, 2, 3, 4}, 2), Result = #table({"First", "FirstLast", "LastFirst", "SkippedFirst", "RangeLast", "EmptyRange", "NoFirst", "NoLast", "DefaultRemove"}, {{List.First(First), List.Last(First), List.First(Last), List.First(Skipped), List.Last(Ranged), List.Count(List.Range({1, 2, 3}, 3)), List.First(NoFirst), List.Last(NoLast), List.First(List.RemoveFirstN({1, 2, 3}))}}) in Result';
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],query),[['First','FirstLast','LastFirst','SkippedFirst','RangeLast','EmptyRange','NoFirst','NoLast','DefaultRemove'],[1,2,3,3,3,0,3,2,2]]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let Values = List.LastN({1, 2, 3}, 0), Result = #table({"Count"}, {{List.Count(Values)}}) in Result'),[['Count'],[0]]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let Values = List.Skip({1, 2, 3}) in #table({"First","Count"}, {{List.First(Values),List.Count(Values)}})'),[['First','Count'],[2,2]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Values = List.Skip({1, 2}, -1), Result = #table({"Count"}, {{List.Count(Values)}}) in Result'),/whole numbers from 0 to 12,000/);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let First = List.FirstN({1, 2, 0, 3}, each _ > 0), Last = List.LastN({0, 2, 3}, each _ > 0), Skipped = List.Skip({1, 2, 0, 3}, each _ > 0), NoFirst = List.RemoveFirstN({1, 2, 0, 3}, each _ > 0), NoLast = List.RemoveLastN({0, 2, 3}, each _ > 0), Result = #table({"FirstLast", "LastFirst", "SkippedFirst", "NoFirstFirst", "NoLastLast"}, {{List.Last(First), List.First(Last), List.First(Skipped), List.First(NoFirst), List.Last(NoLast)}}) in Result'),[['FirstLast','LastFirst','SkippedFirst','NoFirstFirst','NoLastLast'],[2,2,0,0,0]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Values = List.FirstN({1, 2}, each _) in #table({"Count"}, {{List.Count(Values)}})'),/condition must return true or false/);
});
test('bounded M List.MatchesAll and List.MatchesAny short-circuit predicate scans',()=>{
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let All = List.MatchesAll({2, 4, 6}, each Number.Mod(_, 2) = 0), Any = List.MatchesAny({1, 3, 4}, each _ = 4), None = List.MatchesAny({1, 3}, each _ = 4) in #table({"All","Any","None"}, {{All,Any,None}})'),[['All','Any','None'],[true,true,false]]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let All = List.MatchesAll({}, each _ > 0), Any = List.MatchesAny({}, each _ > 0) in #table({"All","Any"}, {{All,Any}})'),[['All','Any'],[true,false]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let All = List.MatchesAll({1}, each _) in #table({"All"}, {{All}})'),/condition must return true or false/);
});
test('bounded M list set operations retain bag counts, stable ordering, nulls and scalar types',()=>{
 const query='let Union = List.Union({{1, 1, 2, null, true}, {1, 2, 2, null, false}}), Intersection = List.Intersect({{1, 1, 2, null, true}, {1, 2, null, null, false}}), Difference = List.Difference({1, 1, 2, 3, null}, {1, 3, 3}), Result = #table({"UnionCount", "UnionFirst", "UnionLast", "IntersectCount", "IntersectLast", "DifferenceCount", "DifferenceFirst", "DifferenceLast", "TypedCount"}, {{List.Count(Union), List.First(Union), List.Last(Union), List.Count(Intersection), List.Last(Intersection), List.Count(Difference), List.First(Difference), List.Last(Difference), List.Count(List.Union({{1, true}, {1, true}}))}}) in Result';
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],query),[['UnionCount','UnionFirst','UnionLast','IntersectCount','IntersectLast','DifferenceCount','DifferenceFirst','DifferenceLast','TypedCount'],[7,1,false,3,null,3,1,null,2]]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let Empty = List.Intersect({}), Union = List.Union({{}, {1}}), Result = #table({"IntersectEmpty", "UnionValue"}, {{List.IsEmpty(Empty), List.First(Union)}}) in Result'),[['IntersectEmpty','UnionValue'],[true,1]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Values = List.Union({{1}, "not a list"}) in #table({"Count"}, {{List.Count(Values)}})'),/expects a list of lists/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Values = List.Intersect({{1}, {1}}, each _ = 1) in #table({"Count"}, {{List.Count(Values)}})'),/expects 1 argument/);
});
test('bounded M list membership helpers support all/any, distinctness and first matching position',()=>{
 const query='let Result = #table({"Distinct", "HasDuplicates", "ContainsAll", "ContainsAny", "ContainsAnyEmpty", "ContainsAllEmpty", "FirstAnyPosition", "MissingAnyPosition"}, {{List.IsDistinct({1, "1", null, true}), List.IsDistinct({1, 1}), List.ContainsAll({1, null, true}, {null, 1}), List.ContainsAny({1, null}, {false, null}), List.ContainsAny({1}, {}), List.ContainsAll({1}, {}), List.PositionOfAny({"a", 2, null, true}, {true, null}), List.PositionOfAny({1, 2}, {3, 4})}}) in Result';
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],query),[['Distinct','HasDuplicates','ContainsAll','ContainsAny','ContainsAnyEmpty','ContainsAllEmpty','FirstAnyPosition','MissingAnyPosition'],[true,false,true,true,false,true,2,-1]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Value = List.ContainsAll({1}, 1) in #table({"Value"}, {{Value}})'),/expects a second list/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Value = List.IsDistinct({{1}}) in #table({"Value"}, {{Value}})'),/compares only finite scalar values/);
});
test('bounded M List.MaxN and List.MinN return ordered top and bottom scalar windows',()=>{
 const query='let High = List.MaxN({3, null, 2, 3, -1}, 3), Low = List.MinN({3, null, 2, 3, -1}, 4), Empty = List.MaxN({1, 2}, 0), Result = #table({"HighFirst", "HighLast", "LowFirst", "LowLast", "Empty"}, {{List.First(High), List.Last(High), List.First(Low), List.Last(Low), List.IsEmpty(Empty)}}) in Result';
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],query),[['HighFirst','HighLast','LowFirst','LowLast','Empty'],[3,2,null,3,true]]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let High = List.MaxN({1, 2, 3, 4}, each _ >= 3), Low = List.MinN({1, 2, 3, 4}, each _ < 3) in #table({"HighFirst","HighLast","LowFirst","LowLast"}, {{List.First(High),List.Last(High),List.First(Low),List.Last(Low)}})'),[['HighFirst','HighLast','LowFirst','LowLast'],[4,3,1,2]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Values = List.MaxN({1, 2}, each _) in #table({"Count"}, {{List.Count(Values)}})'),/condition must return true or false/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Values = List.MinN({1, "a"}, 1) in #table({"Count"}, {{List.Count(Values)}})'),/one scalar type/);
});
test('bounded M List.RemoveMatchingItems removes every typed scalar match',()=>{
 const query='let Values = List.RemoveMatchingItems({1, 2, 1, null, "1", true, 1}, {1, null}), Result = #table({"Count", "First", "Last"}, {{List.Count(Values), List.First(Values), List.Last(Values)}}) in Result';
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],query),[['Count','First','Last'],[3,2,true]]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let Values = List.RemoveMatchingItems({1, 2}, {}), Result = #table({"Count"}, {{List.Count(Values)}}) in Result'),[['Count'],[2]]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let Values = List.RemoveItems({1, 2, 2, 3}, {2}), Result = #table({"First", "Count"}, {{List.First(Values), List.Count(Values)}}) in Result'),[['First','Count'],[1,2]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Values = List.RemoveMatchingItems({1}, {1}, each _) in #table({"Count"}, {{List.Count(Values)}})'),/expects 2 arguments/);
});
test('bounded M List.Single and List.SingleOrDefault validate cardinality and defaults',()=>{
 const query='let Result = #table({"One", "OneNull", "Default", "SingleOrDefault", "SinglePivot"}, {{List.Single({7}), List.Single({null}), List.SingleOrDefault({}, -1), List.SingleOrDefault({"only"}), List.First(Table.Column(Table.Pivot(#table({"Group", "Kind", "Value"}, {{"A", "x", 5}}), {"x"}, "Kind", "Value", List.Single), "x"))}}) in Result';
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],query),[['One','OneNull','Default','SingleOrDefault','SinglePivot'],[7,null,-1,'only',5]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Value = List.Single({}) in #table({"Value"}, {{Value}})'),/exactly one list item/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Value = List.SingleOrDefault({1, 2}, 0) in #table({"Value"}, {{Value}})'),/multiple items/);
});
test('bounded M List.ReplaceMatchingItems applies ordered typed replacement pairs',()=>{
 const query='let Values = List.ReplaceMatchingItems({1, "1", 2, 1, null}, {{1, 2}, {2, 9}, {null, "blank"}}), Result = #table({"Count", "First", "TextNumber", "Last"}, {{List.Count(Values), List.First(Values), List.First(List.Range(Values, 1, 1)), List.Last(Values)}}) in Result';
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],query),[['Count','First','TextNumber','Last'],[5,9,'1','blank']]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Values = List.ReplaceMatchingItems({1}, {{1}}) in #table({"Count"}, {{List.Count(Values)}})'),/oldValue, newValue/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Values = List.ReplaceMatchingItems({1}, {{1, {2}}}) in #table({"Count"}, {{List.Count(Values)}})'),/finite scalar values/);
});
test('bounded M converts scalar tables to records and reconstructs records with explicit missing-field behavior',()=>{
 const roundtrip='let Records = Table.ToRecords(Source), Rebuilt = Table.FromRecords(Records) in Rebuilt';
 const input=[['Name','Count'],['Ada',3],['Lin',null]];
 assert.deepEqual(executePowerQueryM(input,roundtrip),input);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let Rows = Table.FromRecords({[Name="Ada", Score=8], [Name="Lin"]}, {"Name", "Score"}, MissingField.UseNull) in Rows'),[['Name','Score'],['Ada',8],['Lin',null]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Rows = Table.FromRecords({[Name="Ada"], [Name="Lin"]}, {"Name", "Score"}) in Rows'),/missing field “Score”/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Rows = Table.FromRecords({[Name={"Ada"}]}) in Rows'),/must contain a scalar value/);
 assert.throws(()=>executePowerQueryM([['Name'],['Ada']],'let Records = Table.ToRecords(Source) in Table.FromRecords(Records, {"Name", "Extra"})'),/missing field “Extra”/);
});
test('bounded M builds tables from scalar rows with generated or explicit headers and validates shape',()=>{
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let Rows = Table.FromRows({{1, "Ada"}, {2, "Lin"}}, {"ID", "Name"}) in Rows'),[['ID','Name'],[1,'Ada'],[2,'Lin']]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let Rows = Table.FromRows({{true, null}, {false, 3}}) in Rows'),[['Column1','Column2'],[true,null],[false,3]]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let Rows = Table.FromRows({}, {"ID", "Name"}) in Rows'),[['ID','Name']]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Rows = Table.FromRows({{1, 2}, {3}}, {"A", "B"}) in Rows'),/row 2 has a different number/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Rows = Table.FromRows({{1, {2}}}, {"A", "B"}) in Rows'),/must contain a scalar value/);
 assert.deepEqual(executePowerQueryM([['Name','Count'],['Ada',3],['Lin',null]],'let Rows = Table.ToRows(Source), Rebuilt = Table.FromRows(Rows, Table.ColumnNames(Source)) in Rebuilt'),[['Name','Count'],['Ada',3],['Lin',null]]);
});
test('bounded M Table.FromValue creates a bounded scalar Value table with a configurable header',()=>{
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let T = Table.FromValue({1,"Bob",null}), Named = Table.FromValue(42,[DefaultColumnName="Answer"]), Result = #table({"Rows","Header","Value"}, {{Table.RowCount(T),List.First(Table.ColumnNames(Named)),List.First(Table.Column(Named,"Answer"))}}) in T'),[['Value'],[1],['Bob'],[null]]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let T = Table.FromValue(42,[DefaultColumnName="Answer"]) in T'),[['Answer'],[42]]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let T = Table.FromValue(42,[]) in T'),[['Value'],[42]]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let T = Table.FromValue({}) in #table({"Rows","Columns"}, {{Table.RowCount(T),Table.ColumnCount(T)}})'),[['Rows','Columns'],[0,1]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let T = Table.FromValue({{1,2}}) in T'),/finite scalar values/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let T = Table.FromValue(1,[Other="No"]) in T'),/only the DefaultColumnName option/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],`let T = Table.FromValue(Text.ToList("${'x'.repeat(1001)}")) in T`),/limited to 1000 rows/);
});
test('bounded M converts tables to columns and pads short input columns with nulls',()=>{
 const input=[['Name','Score'],['Ada',8],['Lin',null]];
 assert.deepEqual(executePowerQueryM(input,'let Columns = Table.ToColumns(Source), Rebuilt = Table.FromColumns(Columns, Table.ColumnNames(Source)) in Rebuilt'),input);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let Built = Table.FromColumns({{"A", "B"}, {1}}, {"Name", "ID"}) in Built'),[['Name','ID'],['A',1],['B',null]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Built = Table.FromColumns({{1}, {2}}, {"Value", "value"}) in Built'),/one unique, non-empty name/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Built = Table.FromColumns({{1, {2}}}) in Built'),/must contain a scalar value/);
});
test('bounded M Table.Combine unions matching headers and rejects case-only header conflicts',()=>{
 const expression='let Combined = Table.Combine({#table({"A", "B"}, {{1, 2}}), #table({"B", "C"}, {{3, 4}})}) in Combined';
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],expression),[['A','B','C'],[1,2,null],[null,3,4]]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let Combined = Table.Combine({#table({"A","Extra"},{{1,"x"}}),#table({"B","A"},{{2,3}})}, {"B","A"}) in Combined'),[['B','A'],[null,1],[2,3]]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let Combined = Table.Combine({}, {"Only"}) in Combined'),[['Only']]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Combined = Table.Combine({#table({"A"},{{1}})}, {"A","a"}) in Combined'),/duplicate names/);
 const hundredRows=Array.from({length:101},()=>'{1}').join(','),manyTables=Array.from({length:10},()=>`#table({"V"},{${hundredRows}})`).join(',');
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],`let Combined = Table.Combine({${manyTables}}) in Combined`),/exceeds the 1,000-row or 12,000-cell limit/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Combined = Table.Combine({#table({"A"}, {{1}}), #table({"a"}, {{2}})}) in Combined'),/differ only by case/);
});
test('bounded M column renaming supports explicit ignore behavior without creating duplicate headers',()=>{
 const input=[['Name','Amount'],['Ada',3]];
 assert.deepEqual(executePowerQueryM(input,'let Renamed = Table.RenameColumns(Source, {{"Name", "Customer"}, {"Missing", "Ignored"}}, MissingField.Ignore) in Renamed'),[['Customer','Amount'],['Ada',3]]);
 assert.throws(()=>executePowerQueryM(input,'let Renamed = Table.RenameColumns(Source, {{"Missing", "Ignored"}}) in Renamed'),/cannot find/);
 assert.throws(()=>executePowerQueryM(input,'let Renamed = Table.RenameColumns(Source, {{"Name", "Amount"}}) in Renamed'),/duplicate columns/);
});
test('bounded M transforms all column names with safe text functions and rejects collisions',()=>{
 const input=[['  first name  ','City'],['Ada','London']];
 assert.deepEqual(executePowerQueryM(input,'let CleanNames = Table.TransformColumnNames(Source, Text.Trim), UpperNames = Table.TransformColumnNames(CleanNames, Text.Upper) in UpperNames'),[['FIRST NAME','CITY'],['Ada','London']]);
 assert.deepEqual(executePowerQueryM(input,'let CleanNames = Table.TransformColumnNames(Source, Text.Trim, []) in CleanNames'),[['first name','City'],['Ada','London']]);
 assert.throws(()=>executePowerQueryM([['ı','I'],['Ada','A']], 'let UpperNames = Table.TransformColumnNames(Source, Text.Upper) in UpperNames'),/unique without case sensitivity/);
 assert.throws(()=>executePowerQueryM(input,'let Names = Table.TransformColumnNames(Source, Text.Length) in Names'),/supports Text.Trim/);
});
test('bounded M date hierarchy helpers expose day, week, month-end and quarter-end levels',()=>{
 const result=executePowerQueryM([['Date'],['2024-02-29']], 'let DayOfYear = Table.AddColumn(Source, "DayOfYear", each Date.DayOfYear([Date]), type number), Week = Table.AddColumn(DayOfYear, "Week", each Date.WeekOfYear([Date], Day.Monday), type number), MonthEnd = Table.AddColumn(Week, "MonthEnd", each Date.EndOfMonth([Date]), type number), QuarterEnd = Table.AddColumn(MonthEnd, "QuarterEnd", each Date.EndOfQuarter([Date]), type number), Weekday = Table.AddColumn(QuarterEnd, "Weekday", each Date.DayOfWeek([Date], Day.Monday), type number), Days = Table.AddColumn(Weekday, "DaysInMonth", each Date.DaysInMonth([Date]), type number) in Days');
 assert.deepEqual(result[0],['Date','DayOfYear','Week','MonthEnd','QuarterEnd','Weekday','DaysInMonth']);
 assert.equal(result[1]?.[1],60);assert.equal(result[1]?.[2],9);assert.equal(result[1]?.[5],3);assert.equal(result[1]?.[6],29);assert.ok(Number(result[1]?.[4])>Number(result[1]?.[3]));
});
test('bounded M adds month and weekday name, month-week and leap-year date helpers',()=>{
 const result=executePowerQueryM([['Date'],['2024-02-29'],['2023-03-15']], 'let Month = Table.AddColumn(Source, "MonthName", each Date.MonthName([Date], "en-US"), type text), Weekday = Table.AddColumn(Month, "WeekdayName", each Date.DayOfWeekName([Date]), type text), Week = Table.AddColumn(Weekday, "WeekOfMonth", each Date.WeekOfMonth([Date], Day.Sunday), type number), Leap = Table.AddColumn(Week, "Leap", each Date.IsLeapYear([Date]), type logical) in Leap');
 assert.deepEqual(result,[['Date','MonthName','WeekdayName','WeekOfMonth','Leap'],['2024-02-29','February','Thursday',5,true],['2023-03-15','March','Wednesday',3,false]]);
 assert.throws(()=>executePowerQueryM([['Date'],['2024-02-29']], 'let Result = Table.AddColumn(Source, "MonthName", each Date.MonthName([Date], "fr-FR"), type text) in Result'),/English cultures only/);
});
test('bounded M #date creates validated dates usable by hierarchy helpers',()=>{
 const result=executePowerQueryM([['Seed'],[0]],'let DateValue = #date(2011, 12, 31), Values = #table({"Serial", "Month", "Weekday", "Leap"}, {{DateValue, Date.MonthName(DateValue, "en-US"), Date.DayOfWeekName(DateValue), Date.IsLeapYear(DateValue)}}) in Values');
 assert.deepEqual(result,[['Serial','Month','Weekday','Leap'],[excelDateSerial('2011-12-31'),'December','Saturday',false]]);
 for(const expression of ['#date(0, 1, 1)','#date(2024, 13, 1)','#date(2023, 2, 29)','#date(2024.5, 1, 1)']) assert.throws(()=>executePowerQueryM([['Seed'],[0]],`let Value = ${expression} in #table({"Value"}, {{Value}})`),/#date requires/);
 assert.throws(()=>executePowerQueryM([['Seed'],[0]],'let Value = #date(2024, 1) in #table({"Value"}, {{Value}})'),/#date expects year, month, and day/);
});
test('bounded M Date.ToRecord exposes calendar parts for table creation',()=>{
 const result=executePowerQueryM([['Seed'],[0]],'let Parts = Date.ToRecord(#date(2024, 2, 29)), Result = Table.FromRecords({Parts}) in Result');
 assert.deepEqual(result,[['Year','Month','Day'],[2024,2,29]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[0]],'let Parts = Date.ToRecord("not a date") in #table({"Value"}, {{1}})'),/valid date serial or ISO date/);
 assert.throws(()=>executePowerQueryM([['Seed'],[0]],'let Parts = Date.ToRecord(null) in #table({"Value"}, {{1}})'),/valid date serial or ISO date/);
});
test('bounded M #datetime and DateTime.ToRecord preserve time-of-day serial fractions',()=>{
 const result=executePowerQueryM([['Seed'],[0]],'let Parts = DateTime.ToRecord(#datetime(2024, 2, 29, 23, 59, 59.5)), Result = Table.FromRecords({Parts}) in Result');
 assert.deepEqual(result,[['Year','Month','Day','Hour','Minute','Second'],[2024,2,29,23,59,59.5]]);
 assert.deepEqual(executePowerQueryM([['Seed'],[0]],'let Parts = DateTime.ToRecord("2024-02-29T12:30:00"), Result = Table.FromRecords({Parts}) in Result'),[['Year','Month','Day','Hour','Minute','Second'],[2024,2,29,12,30,0]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[0]],'let Value = #datetime(2024, 2, 29, 24, 0, 0) in #table({"Value"}, {{Value}})'),/#datetime needs/);
 assert.throws(()=>executePowerQueryM([['Seed'],[0]],'let Value = #datetime(2024, 2, 29, 23) in #table({"Value"}, {{Value}})'),/#datetime expects year, month, day, hour, minute, and second/);
});
test('bounded M DateTime.FromText parses strict local ISO timestamps without guessing zones or locale',()=>{
 const result=executePowerQueryM([['Seed'],[0]],'let Value = DateTime.FromText("2024-02-29T23:59:59.5"), Parts = DateTime.ToRecord(Value), Result = Table.FromRecords({Parts}) in Result');
 assert.deepEqual(result,[['Year','Month','Day','Hour','Minute','Second'],[2024,2,29,23,59,59.5]]);
 assert.deepEqual(executePowerQueryM([['Seed'],[0]],'let Value = DateTime.FromText(null) in #table({"Value"}, {{Value}})'),[['Value'],[null]]);
 for(const value of ['2023-02-29T12:00:00','2024-02-29 12:00:00','2024-02-29T24:00:00','2024-02-29T12:00:00Z','02/29/2024 12:00:00'])
  assert.throws(()=>executePowerQueryM([['Seed'],[0]],`let Value = DateTime.FromText("${value}") in #table({"Value"}, {{Value}})`),/exact ISO local timestamps/);
});
test('bounded M DateTime.ToText emits invariant ISO local timestamps',()=>{
 const result=executePowerQueryM([['Seed'],[0]],'let A = DateTime.ToText(DateTime.FromText("2024-02-29T23:59:59.5")), B = DateTime.ToText(#date(2024, 1, 2)), Result = #table({"Timestamp","Midnight"}, {{A,B}}) in Result');
 assert.deepEqual(result,[['Timestamp','Midnight'],['2024-02-29T23:59:59.5','2024-01-02T00:00:00']]);
 assert.deepEqual(executePowerQueryM([['Seed'],[0]],'let Value = DateTime.ToText(null) in #table({"Value"}, {{Value}})'),[['Value'],[null]]);
});
test('bounded M DateTime.From accepts GridForge serials and strict ISO local values',()=>{
 const result=executePowerQueryM([['Seed'],[0]],'let A = DateTime.From(45292.5), B = DateTime.From("2024-02-29T12:30:00"), C = DateTime.From("2024-02-29"), Result = #table({"Serial","Timestamp","Date"}, {{A,B,C}}) in Result');
 assert.deepEqual(result,[['Serial','Timestamp','Date'],[45292.5,excelDateSerial('2024-02-29')+(12*3600+30*60)/86400,excelDateSerial('2024-02-29')]]);
 assert.deepEqual(executePowerQueryM([['Seed'],[0]],'let Value = DateTime.From(null) in #table({"Value"}, {{Value}})'),[['Value'],[null]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[0]],'let Value = DateTime.From("02/29/2024") in #table({"Value"}, {{Value}})'),/bounded worksheet serial or ISO local date\/time/);
 assert.throws(()=>executePowerQueryM([['Seed'],[0]],'let Value = DateTime.From("2024-02-29T12:30:00Z") in #table({"Value"}, {{Value}})'),/bounded worksheet serial or ISO local date\/time/);
});
test('bounded M DateTime offsets preserve fractional days across calendar boundaries',()=>{
 const result=executePowerQueryM([['Seed'],[0]],'let Start = DateTime.FromText("2024-02-29T23:30:00"), Shifted = DateTime.AddHours(Start, 2), Parts = DateTime.ToRecord(Shifted), Result = Table.FromRecords({Parts}) in Result');
 assert.deepEqual(result,[['Year','Month','Day','Hour','Minute','Second'],[2024,3,1,1,30,0]]);
 const seconds=executePowerQueryM([['Seed'],[0]],'let Start = DateTime.FromText("2024-01-01T00:00:00"), Shifted = DateTime.AddSeconds(Start, 1.5), Parts = DateTime.ToRecord(Shifted), Result = Table.FromRecords({Parts}) in Result');
 assert.deepEqual(seconds,[['Year','Month','Day','Hour','Minute','Second'],[2024,1,1,0,0,1.5]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[0]],'let Value = DateTime.AddDays(0, -1) in #table({"Value"}, {{Value}})'),/outside the supported Excel datetime range/);
 assert.deepEqual(executePowerQueryM([['Seed'],[0]],'let Value = DateTime.AddHours(null, 1) in #table({"Value"}, {{Value}})'),[['Value'],[null]]);
});
test('bounded M DateTime.Date and DateTime.Time split a timestamp into serial day and fraction',()=>{
 const result=executePowerQueryM([['Seed'],[0]],'let Stamp = DateTime.FromText("2024-05-06T12:30:00"), Day = DateTime.Date(Stamp), Time = DateTime.Time(Stamp), Result = #table({"Day","Time"}, {{Day,Time}}) in Result');
 assert.equal(result[1][0],excelDateSerial('2024-05-06'));
 assert.ok(Math.abs(result[1][1]-0.5208333333333334)<1e-9);
 assert.deepEqual(executePowerQueryM([['Seed'],[0]],'let Result = #table({"DateOnly","TimeOnly"}, {{DateTime.Date("2024-05-06"), DateTime.Time("2024-05-06")}}) in Result'),[['DateOnly','TimeOnly'],[excelDateSerial('2024-05-06'),0]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[0]],'let Value = DateTime.Time("2024-05-06T12:00:00Z") in #table({"Value"}, {{Value}})'),/valid worksheet serial or ISO local date\/time/);
});
test('bounded M DateTime start/end-of-day and calendar offsets retain the time component',()=>{
 const result=executePowerQueryM([['Seed'],[0]],'let Stamp = DateTime.FromText("2024-02-29T12:30:00"), Shifted = DateTime.AddYears(Stamp, 1), Parts = DateTime.ToRecord(Shifted), End = DateTime.ToRecord(DateTime.EndOfDay(Stamp)), Result = Table.FromRecords({Parts, End}) in Result');
 assert.deepEqual(result,[['Year','Month','Day','Hour','Minute','Second'],[2025,2,28,12,30,0],[2024,2,29,23,59,59]]);
 const day=executePowerQueryM([['Seed'],[0]],'let Stamp = DateTime.FromText("2024-02-29T12:30:00"), StartRecord = DateTime.ToRecord(DateTime.StartOfDay(Stamp)), EndRecord = DateTime.ToRecord(DateTime.EndOfDay(Stamp)), Result = #table({"Start","End"}, {{StartRecord[Hour], EndRecord[Hour]}}) in Result');
 assert.deepEqual(day,[['Start','End'],[0,23]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[0]],'let Stamp = DateTime.FromText("2024-02-29T12:30:00"), Shifted = DateTime.AddMonths(Stamp, 1.5) in #table({"Value"}, {{Shifted}})'),/month, quarter and year offsets must be whole numbers/);
});
test('bounded M Time constructor, parser and extractors use worksheet day fractions',()=>{
 const parts=executePowerQueryM([['Seed'],[0]],'let A = Time.ToRecord(#time(23, 59, 59.5)), B = Time.ToRecord(Time.FromText("01:02:03.25")), Result = Table.FromRecords({A,B}) in Result');
 assert.deepEqual(parts,[['Hour','Minute','Second'],[23,59,59.5],[1,2,3.25]]);
 assert.deepEqual(executePowerQueryM([['Seed'],[0]],'let Value = Time.Hour(Time.FromText("08:09:10")) in #table({"Hour"}, {{Value}})'),[['Hour'],[8]]);
 for(const value of ['24:00:00','12:60:00','12:00:60','12:00:00Z','12:00:00.'])
  assert.throws(()=>executePowerQueryM([['Seed'],[0]],`let Value = Time.FromText("${value}") in #table({"Value"}, {{Value}})`),/exact local text/);
 assert.throws(()=>executePowerQueryM([['Seed'],[0]],'let Value = #time(24, 0, 0) in #table({"Value"}, {{Value}})'),/#time requires/);
});
test('bounded M Time.From accepts time fractions and strict local time text',()=>{
 const result=executePowerQueryM([['Seed'],[0]],'let A = Time.From(0.5), B = Time.From("06:30:00"), C = Time.From(null), Result = #table({"A","B","C"}, {{A,B,C}}) in Result');
 assert.deepEqual(result,[['A','B','C'],[0.5,0.2708333333333333,null]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[0]],'let Value = Time.From(1) in #table({"Value"}, {{Value}})'),/Time.From accepts/);
 assert.throws(()=>executePowerQueryM([['Seed'],[0]],'let Value = Time.From("06:30:00", "en-US") in #table({"Value"}, {{Value}})'),/culture parsing is unsupported/);
});
test('bounded M Time.StartOfHour and Time.EndOfHour return hour boundaries',()=>{
 const result=executePowerQueryM([['Seed'],[0]],'let Value = Time.FromText("06:30:00"), Start = Time.ToRecord(Time.StartOfHour(Value)), End = Time.ToRecord(Time.EndOfHour(Value)), Result = Table.FromRecords({Start,End}) in Result');
 assert.deepEqual(result,[['Hour','Minute','Second'],[6,0,0],[6,59,59]]);
});
test('bounded M Time.ToText emits invariant 24-hour text and millisecond fractions',()=>{
 const result=executePowerQueryM([['Seed'],[0]],'let A = Time.ToText(#time(6, 7, 8)), B = Time.ToText(#time(23, 59, 59.125)), Result = #table({"Whole","Fractional"}, {{A,B}}) in Result');
 assert.deepEqual(result,[['Whole','Fractional'],['06:07:08','23:59:59.125']]);
 assert.deepEqual(executePowerQueryM([['Seed'],[0]],'let Value = Time.ToText(null) in #table({"Time"}, {{Value}})'),[['Time'],[null]]);
});
test('bounded M Duration.Total functions expose checked elapsed units',()=>{
 const result=executePowerQueryM([['Seed'],[0]],'let Span = #duration(1, 12, 0, 0), Result = #table({"Days","Hours","Minutes","Seconds"}, {{Duration.TotalDays(Span), Duration.TotalHours(Span), Duration.TotalMinutes(Span), Duration.TotalSeconds(Span)}}) in Result');
 assert.deepEqual(result,[['Days','Hours','Minutes','Seconds'],[1.5,36,2160,129600]]);
 assert.deepEqual(executePowerQueryM([['Seed'],[0]],'let Value = Duration.TotalSeconds(null) in #table({"Value"}, {{Value}})'),[['Value'],[null]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[0]],'let Value = Duration.TotalDays(1) in #table({"Value"}, {{Value}})'),/expects a duration value/);
 assert.throws(()=>executePowerQueryM([['Seed'],[0]],'let Value = Duration.TotalHours(#duration(1e308, 0, 0, 0)) in #table({"Value"}, {{Value}})'),/finite numeric range/);
});
test('bounded M Duration.From and component accessors normalize strict elapsed values',()=>{
 const result=executePowerQueryM([['Seed'],[0]],'let A = Duration.ToRecord(Duration.FromText("1.05:55:20.25")), B = Duration.ToRecord(Duration.From(-0.5)), Result = Table.FromRecords({A,B}) in Result');
 assert.deepEqual(result,[['Days','Hours','Minutes','Seconds'],[1,5,55,20.25],[0,-12,0,0]]);
 assert.equal(executePowerQueryM([['Seed'],[0]],'let Value = Duration.Seconds(Duration.FromText("00:01:02.5")) in #table({"Seconds"}, {{Value}})')[1][0],2.5);
 assert.equal(executePowerQueryM([['Seed'],[0]],'let Value = Duration.TotalHours(Duration.From(#duration(1, 12, 0, 0))) in #table({"Hours"}, {{Value}})')[1][0],36);
 assert.equal(executePowerQueryM([['Seed'],[0]],'let Value = Duration.TotalMinutes(Duration.FromText("12:30")) in #table({"Minutes"}, {{Value}})')[1][0],750);
 assert.deepEqual(executePowerQueryM([['Seed'],[0]],'let Value = Duration.FromText(null) in #table({"Value"}, {{Value}})'),[['Value'],[null]]);
 for(const value of ['24:00:00','1.24:00:00','00:60:00','text']) assert.throws(()=>executePowerQueryM([['Seed'],[0]],`let Value = Duration.FromText("${value}") in #table({"Value"}, {{Value}})`),/strict d\.hh:mm:ss duration text/);
 assert.throws(()=>executePowerQueryM([['Seed'],[0]],'let Value = Duration.From("1.00:00:00", "en-US") in #table({"Value"}, {{Value}})'),/culture parsing is unsupported/);
});
test('bounded M Duration.ToText writes normalized elapsed time and rejects legacy format overrides',()=>{
 const result=executePowerQueryM([['Seed'],[0]],'let A = Duration.ToText(#duration(2, 5, 55, 20.25)), B = Duration.ToText(Duration.From(-0.5)), Result = #table({"Positive","Negative"}, {{A,B}}) in Result');
 assert.deepEqual(result,[['Positive','Negative'],['2.05:55:20.25','-0.12:00:00']]);
 assert.deepEqual(executePowerQueryM([['Seed'],[0]],'let Value = Duration.ToText(null) in #table({"Value"}, {{Value}})'),[['Value'],[null]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[0]],'let Value = Duration.ToText(#duration(1, 0, 0, 0), "x") in #table({"Value"}, {{Value}})'),/format is deprecated/);
});
test('bounded M duration arithmetic scales and combines spans and shifts worksheet serials',()=>{
 const result=executePowerQueryM([['Seed'],[0]],'let Span = #duration(1, 12, 0, 0), Combined = Span + #duration(0, 12, 0, 0), Scaled = Span * 2, Half = Span / 2, Negative = -Span, DateTime = #datetime(2024, 1, 1, 12, 0, 0) + #duration(0, 2, 0, 0), Parts = DateTime.ToRecord(DateTime), Result = #table({"CombinedHours","ScaledHours","HalfHours","NegativeHours","ShiftedHour","Greater"}, {{Duration.TotalHours(Combined), Duration.TotalHours(Scaled), Duration.TotalHours(Half), Duration.TotalHours(Negative), Parts[Hour], Span > #duration(1, 0, 0, 0)}}) in Result');
 assert.deepEqual(result,[['CombinedHours','ScaledHours','HalfHours','NegativeHours','ShiftedHour','Greater'],[48,72,18,-36,14,true]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[0]],'let Value = #duration(0, 1, 0, 0) / 0 in #table({"Value"}, {{Value}})'),/non-zero finite number/);
 assert.throws(()=>executePowerQueryM([['Seed'],[0]],'let Value = #duration(0, 1, 0, 0) - 1 in #table({"Value"}, {{Value}})'),/Subtracting a worksheet serial/);
});
test('bounded M List.Dates generates worksheet calendar series from bounded durations',()=>{
 const result=executePowerQueryM([['Seed'],[0]],'let Dates = List.Dates(#date(2024, 1, 1), 3, #duration(1, 0, 0, 0)), Result = Table.FromColumns({Dates}, {"Date"}) in Result');
 assert.deepEqual(result,[['Date'],[excelDateSerial('2024-01-01')],[excelDateSerial('2024-01-02')],[excelDateSerial('2024-01-03')]]);
 assert.deepEqual(executePowerQueryM([['Seed'],[0]],'let Dates = List.Dates(#date(2024, 1, 1), 0, #duration(1, 0, 0, 0)), Result = Table.FromColumns({Dates}, {"Date"}) in Result'),[['Date']]);
 assert.throws(()=>executePowerQueryM([['Seed'],[0]],'let Dates = List.Dates(#date(2024, 1, 1), 12001, #duration(1, 0, 0, 0)) in #table({"Value"}, {{List.Count(Dates)}})'),/count from 0 to 12,000/);
 assert.throws(()=>executePowerQueryM([['Seed'],[0]],'let Dates = List.Dates(#date(2024, 1, 1), 2, 1) in #table({"Value"}, {{List.Count(Dates)}})'),/a #duration step/);
 assert.throws(()=>executePowerQueryM([['Seed'],[0]],'let Step = #duration(1, null, 0, 0) in #table({"Value"}, {{1}})'),/parts must be finite numbers/);
});
test('bounded M date offsets add days, months, quarters and years with month-end clamping',()=>{
 const result=executePowerQueryM([['Date'],['2024-01-31'],['2024-02-29']], 'let Days = Table.AddColumn(Source, "Days", each Date.AddDays([Date], -1), type number), Months = Table.AddColumn(Days, "Months", each Date.AddMonths([Date], 1), type number), Quarters = Table.AddColumn(Months, "Quarters", each Date.AddQuarters([Date], 1), type number), Years = Table.AddColumn(Quarters, "Years", each Date.AddYears([Date], 1), type number) in Years');
 assert.deepEqual(result[0],['Date','Days','Months','Quarters','Years']);
 assert.deepEqual(result.slice(1),[['2024-01-31',excelDateSerial('2024-01-30'),excelDateSerial('2024-02-29'),excelDateSerial('2024-04-30'),excelDateSerial('2025-01-31')],['2024-02-29',excelDateSerial('2024-02-28'),excelDateSerial('2024-03-29'),excelDateSerial('2024-05-29'),excelDateSerial('2025-02-28')]]);
 assert.throws(()=>executePowerQueryM([['Date'],['2024-01-31']], 'let Shifted = Table.AddColumn(Source, "Next", each Date.AddMonths([Date], 1.5), type number) in Shifted'),/whole-number offset/);
});
test('bounded M Date.AddWeeks offsets by exact whole seven-day units',()=>{
 const result=executePowerQueryM([['Date'],['2024-01-31']], 'let Forward = Table.AddColumn(Source, "Forward", each Date.AddWeeks([Date], 1), type number), Back = Table.AddColumn(Forward, "Back", each Date.AddWeeks([Date], -1), type number) in Back');
 assert.deepEqual(result,[['Date','Forward','Back'],['2024-01-31',excelDateSerial('2024-02-07'),excelDateSerial('2024-01-24')]]);
 assert.throws(()=>executePowerQueryM([['Date'],['2024-01-31']], 'let Result = Table.AddColumn(Source, "Shift", each Date.AddWeeks([Date], 1.5), type number) in Result'),/whole-number offset/);
});
test('bounded M Date.From and Date.FromText parse ISO dates without locale guessing',()=>{
 const result=executePowerQueryM([['Text','Serial'],['2024-02-29',excelDateSerial('2024-02-29')+0.75]], 'let Parsed = Table.AddColumn(Source, "Parsed", each Date.FromText([Text]), type date), Normalized = Table.AddColumn(Parsed, "DateValue", each Date.From([Serial]), type date) in Normalized');
 assert.deepEqual(result[0],['Text','Serial','Parsed','DateValue']);assert.equal(result[1]?.[2],excelDateSerial('2024-02-29'));assert.equal(result[1]?.[3],excelDateSerial('2024-02-29'));
 assert.equal(executePowerQueryM([['Seed'],[0]],'let Value = Date.From("2024-02-29T12:30:00") in #table({"Date"}, {{Value}})')[1][0],excelDateSerial('2024-02-29'));
 assert.throws(()=>executePowerQueryM([['Seed'],[0]],'let Value = Date.From("2024-02-29T12:30:00Z") in #table({"Date"}, {{Value}})'),/ISO local date\/time text/);
 assert.throws(()=>executePowerQueryM([['Text'],['03/04/2024']], 'let Parsed = Table.AddColumn(Source, "Parsed", each Date.FromText([Text]), type date) in Parsed'),/exact ISO date text/);
 assert.throws(()=>executePowerQueryM([['Value'],['2024-02-29']], 'let Parsed = Table.AddColumn(Source, "Parsed", each Date.FromText(42), type date) in Parsed'),/accepts ISO date text only/);
});
test('bounded M week boundary helpers respect selected week starts and Excel serial 60',()=>{
 const result=executePowerQueryM([['Date'],['2024-02-29']], 'let Start = Table.AddColumn(Source, "WeekStart", each Date.StartOfWeek([Date], Day.Monday), type number), End = Table.AddColumn(Start, "WeekEnd", each Date.EndOfWeek([Date], Day.Monday), type number) in End');
 assert.equal(result[1]?.[1],excelDateSerial('2024-02-26'));assert.equal(result[1]?.[2],excelDateSerial('2024-03-03'));
 const leapDay=executePowerQueryM([['Date'],['1900-02-29']], 'let DOW = Table.AddColumn(Source, "Weekday", each Date.DayOfWeek([Date], Day.Monday), type number) in DOW');
 assert.equal(leapDay[1]?.[1],3);
});
test('bounded M nested joins expand lookup columns and keep unmatched left rows',()=>{
 const expression='let Lookup = #table({"ID", "City"}, {{1, "West"}, {2, "East"}}), Joined = Table.NestedJoin(Source, {"ID"}, Lookup, {"ID"}, "Matches", JoinKind.LeftOuter), Expanded = Table.ExpandTableColumn(Joined, "Matches", {"City"}) in Expanded';
 assert.deepEqual(executePowerQueryM([['ID','Name'],[1,'Ada'],[3,'Lin']],expression),[['ID','Name','City'],[1,'Ada','West'],[3,'Lin',null]]);
 assert.deepEqual(executePowerQueryM([['ID','Name']],expression),[['ID','Name','City']]);
});
test('bounded M Table.DuplicateColumn copies values and preserves nested join schemas',()=>{
 const input=[['Name','Score'],['Ada',8],['Lin',null]];
 assert.deepEqual(executePowerQueryM(input,'let Duplicated = Table.DuplicateColumn(Source, "Score", "Score copy") in Duplicated'),[['Name','Score','Score copy'],['Ada',8,8],['Lin',null,null]]);
 assert.throws(()=>executePowerQueryM(input,'let Duplicated = Table.DuplicateColumn(Source, "Score", "name") in Duplicated'),/unique non-empty output name/);
 assert.throws(()=>executePowerQueryM(input,'let Duplicated = Table.DuplicateColumn(Source, "Missing", "Copy") in Duplicated'),/existing source/);
 const nested='let Lookup = #table({"ID", "City"}, {{1, "Paris"}}), Joined = Table.NestedJoin(Source, {"ID"}, Lookup, {"ID"}, "Matches", JoinKind.LeftOuter), Duplicated = Table.DuplicateColumn(Joined, "Matches", "MatchesCopy"), Expanded = Table.ExpandTableColumn(Duplicated, "Matches", {"City"}, {"City"}), ExpandedCopy = Table.ExpandTableColumn(Expanded, "MatchesCopy", {"City"}, {"CopyCity"}) in ExpandedCopy';
 assert.deepEqual(executePowerQueryM([['ID'],[1]],nested),[['ID','City','CopyCity'],[1,'Paris','Paris']]);
 assert.throws(()=>executePowerQueryM(input,'let Duplicated = Table.DuplicateColumn(Source, "Score", "Copy", type number) in Duplicated'),/expects 3 arguments/);
});
test('bounded M Table.SplitColumn expands text with an explicit delimiter splitter',()=>{
 const expression='let Splitter = Splitter.SplitTextByDelimiter("|", QuoteStyle.None), Split = Table.SplitColumn(Source, "Location", Splitter, {"City", "State"}) in Split';
 assert.deepEqual(executePowerQueryM([['ID','Location'],[1,'Paris|TX'],[2,'London'],[3,null]],expression),[['ID','City','State'],[1,'Paris','TX'],[2,'London',null],[3,null,null]]);
 assert.throws(()=>executePowerQueryM([['Location'],['a|b|c']],expression),/pieces but only 2 output names/);
 const csv='let Splitter = Splitter.SplitTextByDelimiter("|", QuoteStyle.Csv), Split = Table.SplitColumn(Source, "Location", Splitter, {"A", "B", "C"}) in Split';
 assert.deepEqual(executePowerQueryM([['Location'],['Paris|"Austin|Round Rock"|TX'],['a|"B ""C"" D"|X']],csv),[['A','B','C'],['Paris','Austin|Round Rock','TX'],['a','B "C" D','X']]);
 assert.throws(()=>executePowerQueryM([['Location'],['a|"b']],csv),/unmatched quote/);
 const joined='let Lookup = #table({"ID", "City"}, {{1, "Paris"}}), Nested = Table.NestedJoin(Source, {"ID"}, Lookup, {"ID"}, "Matches", JoinKind.LeftOuter), Splitter = Splitter.SplitTextByDelimiter("|"), Split = Table.SplitColumn(Nested, "Location", Splitter, {"City", "State"}), Expanded = Table.ExpandTableColumn(Split, "Matches", {"City"}, {"LookupCity"}) in Expanded';
 assert.deepEqual(executePowerQueryM([['ID','Location'],[1,'Paris|TX']],joined),[['ID','City','State','LookupCity'],[1,'Paris','TX','Paris']]);
});
test('bounded M Table.CombineColumns joins text, skips nulls, and quotes CSV fields when required',()=>{
 const plain='let Combined = Table.CombineColumns(Source, {"Last", "First"}, Combiner.CombineTextByDelimiter(", "), "Name") in Combined';
 assert.deepEqual(executePowerQueryM([['ID','First','Middle','Last'],[1,'Ada',null,'Lovelace'],[2,'Lin','Q','']],plain),[['ID','Name','Middle'],[1,'Lovelace, Ada',null],[2,', Lin','Q']]);
 const csv='let Combined = Table.CombineColumns(Source, {"A", "B", "C"}, Combiner.CombineTextByDelimiter("|", QuoteStyle.Csv), "Merged") in Combined';
 assert.deepEqual(executePowerQueryM([['A','B','C'],['north','Austin|Round Rock','TX'],['a','B "C" D','x'],[null,'solo',null]],csv),[['Merged'],['north|"Austin|Round Rock"|TX'],['a|"B ""C"" D"|x'],['solo']]);
 const roundtrip='let Combined = Table.CombineColumns(Source, {"Left", "Right"}, Combiner.CombineTextByDelimiter(",", QuoteStyle.Csv), "Merged"), Splitter = Splitter.SplitTextByDelimiter(",", QuoteStyle.Csv), Split = Table.SplitColumn(Combined, "Merged", Splitter, {"LeftOut", "RightOut"}) in Split';
 assert.deepEqual(executePowerQueryM([['Left','Right'],['line\nbreak','plain'],['a"b',null]],roundtrip),[['LeftOut','RightOut'],['line\nbreak','plain'],['a"b',null]]);
 assert.throws(()=>executePowerQueryM([['ID','First','Middle','Last'],[1,1,null,'x']],plain),/must be text or null/);
 assert.throws(()=>executePowerQueryM([['A','B','C'],['a','b','kept']],'let Combined = Table.CombineColumns(Source, {"A", "B"}, Combiner.CombineTextByDelimiter("|"), "C") in Combined'),/conflicts/);
 assert.throws(()=>executePowerQueryM([['A','B'],['x'.repeat(32767),'y']],'let Combined = Table.CombineColumns(Source, {"A", "B"}, Combiner.CombineTextByDelimiter(""), "Merged") in Combined'),/32,767-character cell limit/);
});
test('bounded M Table.Sort orders nulls, logicals, numbers, and text consistently and stays stable',()=>{
 const input=[['Value','Tie'],['b',1],[10,1],[null,1],[true,1],[2,1],[false,1],['a',1],[2,0]];
 const sorted=executePowerQueryM(input,'let Sorted = Table.Sort(Source, {{"Value", Order.Ascending}}) in Sorted');
 assert.deepEqual(sorted,[['Value','Tie'],[null,1],[false,1],[true,1],[2,1],[2,0],[10,1],['a',1],['b',1]]);
 const stable=executePowerQueryM([['Key','Order'],['x',1],['x',2],['y',3]],'let Sorted = Table.Sort(Source, {{"Key", Order.Ascending}}) in Sorted');
 assert.deepEqual(stable,[['Key','Order'],['x',1],['x',2],['y',3]]);
 assert.deepEqual(executePowerQueryM([['V'],[1],[null],[2]],'let Sorted = Table.Sort(Source, {{"V", Order.Descending}}) in Sorted'),[['V'],[2],[1],[null]]);
});
test('bounded M Table.Distinct preserves first rows and type-aware composite keys',()=>{
 const input=[['Region','Code','Keep'],['East',1,'first'],['East',1,'duplicate'],['East','1','text key'],['West',null,'blank'],['West',null,'duplicate blank']];
 assert.deepEqual(executePowerQueryM(input,'let Unique = Table.Distinct(Source, {"Region", "Code"}) in Unique'),[['Region','Code','Keep'],['East',1,'first'],['East','1','text key'],['West',null,'blank']]);
 assert.deepEqual(executePowerQueryM(input,'let Unique = Table.Distinct(Source, "Region") in Unique'),[['Region','Code','Keep'],['East',1,'first'],['West',null,'blank']]);
 assert.deepEqual(executePowerQueryM([['K','V'],['a',1],['a',2]],'let Unique = Table.Distinct(Source, {}) in Unique'),[['K','V'],['a',1]]);
 assert.throws(()=>executePowerQueryM(input,'let Unique = Table.Distinct(Source, {"Region", "Region"}) in Unique'),/key columns must be unique/);
 assert.throws(()=>executePowerQueryM(input,'let Unique = Table.Distinct(Source, {"Missing"}) in Unique'),/Unknown distinct key column/);
 assert.deepEqual(executePowerQueryM(input,'let Results = #table({"WholeTable", "RegionKey", "PairKey"}, {{Table.IsDistinct(Source), Table.IsDistinct(Source, "Region"), Table.IsDistinct(Source, {"Region", "Code"})}}) in Results'),[['WholeTable','RegionKey','PairKey'],[true,false,false]]);
 assert.deepEqual(executePowerQueryM([['K'],['a']],'let Unique = Table.IsDistinct(Source) in #table({"Unique"}, {{Unique}})'),[['Unique'],[true]]);
});
test('bounded M Table.Pivot aggregates scalars and rejects list-valued aggregate results',()=>{
 const input=[['Team','Kind','Amount'],['A','x',2],['A','x',3],['A','y',4],['B','y',5]];
 assert.deepEqual(executePowerQueryM(input,'let Pivoted = Table.Pivot(Source, {"x", "y"}, "Kind", "Amount", List.Sum) in Pivoted'),[['Team','x','y'],['A',5,4],['B',null,5]]);
 assert.throws(()=>executePowerQueryM(input,'let Pivoted = Table.Pivot(Source, {"x"}, "Kind", "Amount", List.Distinct) in Pivoted'),/must return scalar values/);
});
test('bounded M TransformColumnTypes converts invariant numbers, logicals and ISO dates',()=>{
 const result=executePowerQueryM([['Amount','Enabled','When'],['12.5','TRUE','2024-02-29']], 'let Typed = Table.TransformColumnTypes(Source, {{"Amount", type number}, {"Enabled", type logical}, {"When", type date}}) in Typed');
 assert.equal(result[1]?.[0],12.5);assert.equal(result[1]?.[1],true);assert.equal(typeof result[1]?.[2],'number');
 const datetime=executePowerQueryM([['When'],['2024-02-29T12:30:00']], 'let Typed = Table.TransformColumnTypes(Source, {{"When", type datetime}}) in Typed');
 assert.equal(datetime[1]?.[0],excelDateSerial('2024-02-29')+(12*3600+30*60)/86400);
 assert.throws(()=>executePowerQueryM([['When'],['2024-02-29T12:30:00Z']], 'let Typed = Table.TransformColumnTypes(Source, {{"When", type datetime}}) in Typed'),/Cannot convert/);
 assert.throws(()=>executePowerQueryM([['Amount'],['1,25']], 'let Typed = Table.TransformColumnTypes(Source, {{"Amount", type number}}) in Typed'),/Cannot convert/);
 assert.throws(()=>executePowerQueryM([['Amount'],['1']], 'let Typed = Table.TransformColumnTypes(Source, {{"Amount", type number}}, "de-DE") in Typed'),/culture-specific/);
});
test('bounded M Number operations cover sign, division, modulo, powers, roots and logarithms',()=>{
 const query='let Results = #table({"Sign", "NullSign", "IntegerDivide", "NegativeDivide", "Mod", "NegativeMod", "Power", "Root", "Exp", "Ln", "Log", "Log10"}, {{Number.Sign(-4), Number.Sign(null), Number.IntegerDivide(8.3, 3), Number.IntegerDivide(-8.3, 3), Number.Mod(10.5, 0.2), Number.Mod(-5, 3), Number.Power(5, 3), Number.Sqrt(81), Number.Exp(0), Number.Ln(1), Number.Log(8, 2), Number.Log10(100)}}) in Results';
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],query),[['Sign','NullSign','IntegerDivide','NegativeDivide','Mod','NegativeMod','Power','Root','Exp','Ln','Log','Log10'],[-1,null,2,-2,0.09999999999999942,-2,125,9,1,0,3,2]]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let Results = #table({"PowerNull", "ModNull", "SqrtZero"}, {{Number.Power(null, 2), Number.Mod(3, null), Number.Sqrt(0)}}) in Results'),[['PowerNull','ModNull','SqrtZero'],[null,null,0]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Value = Number.Mod(1, 0) in #table({"Value"}, {{Value}})'),/cannot divide by zero/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Value = Number.Sqrt(-1) in #table({"Value"}, {{Value}})'),/non-negative/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Value = Number.Log(5, 1) in #table({"Value"}, {{Value}})'),/positive base other than 1/);
});
test('bounded M factorial, combinations and permutations validate exact integer domains',()=>{
 const query='let Result = #table({"Factorial", "ZeroFactorial", "Combinations", "Permutation", "EdgeCombination", "NullFactorial", "NullPermutation"}, {{Number.Factorial(10), Number.Factorial(0), Number.Combinations(5, 3), Number.Permutations(5, 3), Number.Combinations(1000000, 1), Number.Factorial(null), Number.Permutations(4, null)}}) in Result';
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],query),[['Factorial','ZeroFactorial','Combinations','Permutation','EdgeCombination','NullFactorial','NullPermutation'],[3628800,1,10,60,1000000,null,null]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Value = Number.Factorial(171) in #table({"Value"}, {{Value}})'),/overflow/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Value = Number.Combinations(4, 5) in #table({"Value"}, {{Value}})'),/cannot exceed/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Value = Number.Permutations(3.5, 1) in #table({"Value"}, {{Value}})'),/whole-number/);
});
test('bounded M Number.From and Number.Round use explicit conversion and rounding rules',()=>{
 assert.deepEqual(executePowerQueryM([['X'],[0]],'let Results = #table({"Even", "Odd", "Away", "Zero"}, {{Number.Round(2.5), Number.Round(3.5), Number.Round(2.5, 0, RoundingMode.AwayFromZero), Number.From(null)}}) in Results'),[['Even','Odd','Away','Zero'],[2,4,3,null]]);
 assert.deepEqual(executePowerQueryM([['X'],[0]],'let Results = #table({"UpPos", "UpNeg", "DownPos", "DownNeg", "TowardZero", "Away", "Even"}, {{Number.Round(1.25, 1, RoundingMode.Up), Number.Round(-1.25, 1, RoundingMode.Up), Number.Round(1.25, 1, RoundingMode.Down), Number.Round(-1.25, 1, RoundingMode.Down), Number.Round(-1.25, 1, RoundingMode.TowardZero), Number.Round(-1.25, 1, RoundingMode.AwayFromZero), Number.Round(-1.25, 1, RoundingMode.ToEven)}}) in Results'),[['UpPos','UpNeg','DownPos','DownNeg','TowardZero','Away','Even'],[1.3,-1.2,1.2,-1.3,-1.2,-1.3,-1.2]]);
 assert.deepEqual(executePowerQueryM([['X'],[0]],'let Results = #table({"Up", "Down", "TowardZero", "Away", "UpDigits", "DownDigits", "Null"}, {{Number.RoundUp(-1.2), Number.RoundDown(-1.2), Number.RoundTowardZero(-1.2), Number.RoundAwayFromZero(-1.2), Number.RoundUp(1.234, 2), Number.RoundDown(-1.234, 2), Number.RoundUp(null)}}) in Results'),[['Up','Down','TowardZero','Away','UpDigits','DownDigits','Null'],[-1,-2,-1,-2,1.24,-1.24,null]]);
 assert.throws(()=>executePowerQueryM([['X'],[0]],'let Results = #table({"N"}, {{Number.From("not a number")}}) in Results'),/could not convert/);
 assert.throws(()=>executePowerQueryM([['Text'],['9']], 'let Kept = Table.SelectRows(Source, each [Text] > 2) in Kept'),/same type/);
});
test('bounded M Number.IsEven and Number.IsOdd classify finite numeric parity',()=>{
 assert.deepEqual(executePowerQueryM([['X'],[0]],'let Results = #table({"EvenPositive", "OddPositive", "EvenNegative", "Fractional", "EvenNull"}, {{Number.IsEven(82), Number.IsOdd(625), Number.IsEven(-4), Number.IsOdd(2.5), Number.IsEven(null)}}) in Results'),[['EvenPositive','OddPositive','EvenNegative','Fractional','EvenNull'],[true,true,true,true,null]]);
 assert.throws(()=>executePowerQueryM([['X'],[0]],'let Result = Number.IsEven("4") in #table({"Value"}, {{Result}})'),/finite number/);
});
test('bounded M trigonometric number helpers validate finite values and inverse domains',()=>{
 const result=executePowerQueryM([['X'],[0]],'let Values = #table({"Sin", "Cos", "Tan", "Asin", "Acos", "Atan", "Atan2", "Sinh", "Cosh", "Tanh", "Null"}, {{Number.Sin(0), Number.Cos(0), Number.Tan(0), Number.Asin(1), Number.Acos(1), Number.Atan(1), Number.Atan2(1, 0), Number.Sinh(0), Number.Cosh(0), Number.Tanh(0), Number.Sin(null)}}) in Values');
 assert.deepEqual(result[0],['Sin','Cos','Tan','Asin','Acos','Atan','Atan2','Sinh','Cosh','Tanh','Null']);
 assert.equal(result[1][0],0);assert.equal(result[1][1],1);assert.equal(result[1][2],0);assert.equal(result[1][3],Math.PI/2);assert.equal(result[1][4],0);assert.equal(result[1][5],Math.PI/4);assert.equal(result[1][6],Math.PI/2);assert.equal(result[1][7],0);assert.equal(result[1][8],1);assert.equal(result[1][9],0);assert.equal(result[1][10],null);
 assert.throws(()=>executePowerQueryM([['X'],[0]],'let Value = Number.Asin(2) in #table({"Value"}, {{Value}})'),/between -1 and 1/);
 assert.throws(()=>executePowerQueryM([['X'],[0]],'let Value = Number.Cosh(1000) in #table({"Value"}, {{Value}})'),/non-finite/);
 assert.throws(()=>executePowerQueryM([['X'],[0]],'let Value = Number.Atan2(1) in #table({"Value"}, {{Value}})'),/expects 2 arguments/);
});
test('bounded M bitwise numeric functions use checked safe-integer operands and shifts',()=>{
 const result=executePowerQueryM([['X'],[0]],'let Values = #table({"And", "Or", "Xor", "Not", "Left", "Right", "NegativeRight", "Null"}, {{Number.BitwiseAnd(6, 3), Number.BitwiseOr(6, 3), Number.BitwiseXor(6, 3), Number.BitwiseNot(5), Number.BitwiseShiftLeft(3, 2), Number.BitwiseShiftRight(13, 2), Number.BitwiseShiftRight(-8, 2), Number.BitwiseAnd(null, 2)}}) in Values');
 assert.deepEqual(result,[['And','Or','Xor','Not','Left','Right','NegativeRight','Null'],[2,7,5,-6,12,3,-2,null]]);
 assert.throws(()=>executePowerQueryM([['X'],[0]],'let Value = Number.BitwiseAnd(1.5, 2) in #table({"Value"}, {{Value}})'),/safe whole-number/);
 assert.throws(()=>executePowerQueryM([['X'],[0]],'let Value = Number.BitwiseShiftLeft(1, 53) in #table({"Value"}, {{Value}})'),/shift count must/);
 assert.throws(()=>executePowerQueryM([['X'],[0]],'let Value = Number.BitwiseShiftLeft(9007199254740991, 1) in #table({"Value"}, {{Value}})'),/safe integer range/);
});
test('bounded M number constants and Number.IsNaN expose special values without loading non-finite cells',()=>{
 const result=executePowerQueryM([['X'],[0]],'let Values = #table({"NaN", "PositiveInfinity", "NegativeInfinity", "Pi", "E", "Epsilon"}, {{Number.IsNaN(Number.NaN), Number.IsNaN(Number.PositiveInfinity), Number.IsNaN(Number.NegativeInfinity), Number.PI, Number.E, Number.Epsilon}}) in Values');
 assert.deepEqual(result[0],['NaN','PositiveInfinity','NegativeInfinity','Pi','E','Epsilon']);
 assert.deepEqual(result[1].slice(0,3),[true,false,false]);
 assert.equal(result[1][3],Math.PI);assert.equal(result[1][4],Math.E);assert.equal(result[1][5],Number.EPSILON);
 assert.throws(()=>executePowerQueryM([['X'],[0]],'let Values = #table({"Value"}, {{Number.PositiveInfinity}}) in Values'),/scalar rows/);
 assert.throws(()=>executePowerQueryM([['X'],[0]],'let Value = Number.IsNaN(null) in #table({"Value"}, {{Value}})'),/expects a number/);
});
test('bounded M Number.FromText parses only invariant decimal text and rejects locale guesses',()=>{
 const result=executePowerQueryM([['X'],[0]],'let Values = #table({"Decimal", "Exponent", "Null", "InvariantCulture"}, {{Number.FromText("-12.50"), Number.FromText("1.25e2"), Number.FromText(null), Number.FromText("4.5", "")}}) in Values');
 assert.deepEqual(result,[['Decimal','Exponent','Null','InvariantCulture'],[-12.5,125,null,4.5]]);
 assert.throws(()=>executePowerQueryM([['X'],[0]],'let Value = Number.FromText("1,234.5") in #table({"Value"}, {{Value}})'),/invariant decimal syntax/);
 assert.throws(()=>executePowerQueryM([['X'],[0]],'let Value = Number.FromText("1,5", "fr-FR") in #table({"Value"}, {{Value}})'),/culture parsing is not implemented/);
 assert.throws(()=>executePowerQueryM([['X'],[0]],'let Value = Number.FromText(4) in #table({"Value"}, {{Value}})'),/expects a text value/);
});
test('bounded M Number.ToText converts numeric values without implying format parity',()=>{
 const result=executePowerQueryM([['X'],[0]],'let Values = #table({"Text", "Exponent", "Null", "NaN", "Infinity"}, {{Number.ToText(4), Number.ToText(1.25e-7), Number.ToText(null), Number.ToText(Number.NaN), Number.ToText(Number.PositiveInfinity)}}) in Values');
 assert.deepEqual(result,[['Text','Exponent','Null','NaN','Infinity'],['4','1.25e-7',null,'NaN','Infinity']]);
 assert.throws(()=>executePowerQueryM([['X'],[0]],'let Value = Number.ToText(4, "N2") in #table({"Value"}, {{Value}})'),/format strings and culture overrides/);
 assert.throws(()=>executePowerQueryM([['X'],[0]],'let Value = Number.ToText("4") in #table({"Value"}, {{Value}})'),/expects a numeric value/);
});
test('bounded M List.AllTrue and List.AnyTrue validate logical lists and empty identities',()=>{
 assert.deepEqual(executePowerQueryM([['A'],[1]],'let Results = #table({"All", "Any", "EmptyAll", "EmptyAny"}, {{List.AllTrue({true, true}), List.AnyTrue({false, true}), List.AllTrue({}), List.AnyTrue({})}}) in Results'),[['All','Any','EmptyAll','EmptyAny'],[true,true,true,false]]);
 assert.throws(()=>executePowerQueryM([['A'],[1]],'let Results = #table({"All"}, {{List.AllTrue({true, 1})}}) in Results'),/logical values/);
});
test('bounded M exposes distinct, empty, contains, and position list helpers through the M parser',()=>{
 const expression='let Values = {1, 1, "1", null, null}, Unique = List.Distinct(Values), Result = #table({"Empty", "HasText", "Missing", "Index", "IndexLast", "IndexMissing", "UniqueCount"}, {{List.IsEmpty({}), List.Contains(Unique, "1"), List.Contains(Unique, 2), List.PositionOf(Unique, null), List.PositionOf({1, 2, 1}, 1, Occurrence.Last), List.PositionOf({1, 2}, 4, Occurrence.Last), List.Count(Unique)}}) in Result';
 assert.deepEqual(executePowerQueryM([['Source'],[1]],expression),[['Empty','HasText','Missing','Index','IndexLast','IndexMissing','UniqueCount'],[true,true,false,2,2,-1,3]]);
 assert.throws(()=>executePowerQueryM([['Source'],[1]],'let Index = List.PositionOf({1, 2}, 1, "last") in #table({"Index"}, {{Index}})'),/Occurrence.First or Occurrence.Last/);
});
test('bounded M list generation and each transforms are deterministic and size-limited',()=>{
 assert.deepEqual(executePowerQueryM([['Value'],[1]],'let Values = List.Numbers(2, 4, 3), Doubled = List.Transform(Values, each _ * 2), Selected = List.Select(Doubled, each _ > 8), Combined = List.Combine({Selected, {null, 20}}), Clean = List.RemoveNulls(Combined), Reversed = List.Reverse(Clean), Rows = Table.FromColumns({Reversed}, {"Value"}) in Rows'),[['Value'],[20],[22],[16],[10]]);
 assert.deepEqual(executePowerQueryM([['Value'],[1]],'let Values = List.Numbers(3, 0), Rows = Table.FromColumns({Values}, {"Value"}) in Rows'),[['Value']]);
 assert.throws(()=>executePowerQueryM([['Value'],[1]],'let Values = List.Numbers(0, 12001), Rows = Table.FromColumns({Values}, {"Value"}) in Rows'),/whole count from 0 to 12,000/);
 assert.throws(()=>executePowerQueryM([['Value'],[1]],'let Values = List.Transform({1, 2}, each {_, _}), Rows = Table.FromColumns({Values}, {"Value"}) in Rows'),/scalar values/);
});
test('bounded M List.Positions, Repeat, Split and Zip handle empty, uneven and capped lists',()=>{
 const result=executePowerQueryM([['Seed'],[1]],'let Positions = List.Positions({"a", "b", "c"}), Repeated = List.Repeat({1, null}, 2), Pages = List.Split({1, 2, 3, 4, 5}, 2), Zipped = List.Zip({{"a", "b"}, {1}, {true, false}}), Result = #table({"PositionLast", "RepeatedCount", "RepeatedNull", "PageCount", "FinalPage", "FinalValue", "ZipRows", "ZipValue", "EmptyRepeat", "EmptyPositions"}, {{List.Last(Positions), List.Count(Repeated), List.Last(Repeated), List.Count(Pages), List.Count(List.Last(Pages)), List.First(List.Last(Pages)), List.Count(Zipped), List.Last(List.Last(Zipped)), List.IsEmpty(List.Repeat({1}, 0)), List.IsEmpty(List.Positions({}))}}) in Result');
 assert.deepEqual(result,[['PositionLast','RepeatedCount','RepeatedNull','PageCount','FinalPage','FinalValue','ZipRows','ZipValue','EmptyRepeat','EmptyPositions'],[2,4,null,3,1,5,2,false,true,true]]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let Rows = List.Zip({{1, 2}, {"a"}}), Result = Table.FromRows(Rows, {"Value", "Name"}) in Result'),[['Value','Name'],[1,'a'],[2,null]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Pages = List.Split({1, 2}, 0) in #table({"Count"}, {{List.Count(Pages)}})'),/positive whole page size/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Values = List.Repeat({1, 2}, 6001) in #table({"Count"}, {{List.Count(Values)}})'),/12,000 output items/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Values = List.Zip({{1}, 2}) in #table({"Count"}, {{List.Count(Values)}})'),/list of lists/);
});
test('bounded M List.Alternate supports skip/keep windows and initial offsets',()=>{
 const query='let SkipFirst = List.Alternate({1, 2, 3, 4, 5}, 1), EveryOther = List.Alternate({1, 2, 3, 4, 5, 6}, 1, 1), OffsetAndPair = List.Alternate({1, 2, 3, 4, 5, 6, 7}, 1, 2, 1), Result = #table({"SkipFirstFirst", "SkipFirstLast", "EveryOtherFirst", "EveryOtherLast", "PairCount", "PairFirst", "PairLast"}, {{List.First(SkipFirst), List.Last(SkipFirst), List.First(EveryOther), List.Last(EveryOther), List.Count(OffsetAndPair), List.First(OffsetAndPair), List.Last(OffsetAndPair)}}) in Result';
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],query),[['SkipFirstFirst','SkipFirstLast','EveryOtherFirst','EveryOtherLast','PairCount','PairFirst','PairLast'],[2,5,2,6,5,1,7]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Values = List.Alternate({1, 2}, 0, 0) in #table({"Count"}, {{List.Count(Values)}})'),/cannot both be zero/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Values = List.Alternate({1, 2}, -1) in #table({"Count"}, {{List.Count(Values)}})'),/whole numbers from 0 to 12,000/);
});
test('bounded M list range edits insert, remove and replace items at zero-based offsets',()=>{
 const query='let Inserted = List.InsertRange({1, 2, 5}, 2, {3, 4}), Removed = List.RemoveRange({1, 2, 3, 4, 5}, 1, 2), Replaced = List.ReplaceRange({1, 2, 7, 8, 9, 5}, 2, 3, {3, 4}), PastEnd = List.RemoveRange({1, 2}, 2, 5), Result = #table({"InsertedCount", "Inserted", "RemovedLast", "RemovedCount", "ReplacedCount", "ReplacedLast", "PastEndCount"}, {{List.Count(Inserted), List.First(List.Range(Inserted, 2, 2)), List.Last(Removed), List.Count(Removed), List.Count(Replaced), List.Last(Replaced), List.Count(PastEnd)}}) in Result';
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],query),[['InsertedCount','Inserted','RemovedLast','RemovedCount','ReplacedCount','ReplacedLast','PastEndCount'],[5,3,5,3,5,5,2]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Values = List.InsertRange({1}, 2, {3}) in #table({"Count"}, {{List.Count(Values)}})'),/valid zero-based index/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Values = List.RemoveRange({1}, 0, -1) in #table({"Count"}, {{List.Count(Values)}})'),/valid zero-based index/);
});
test('bounded M Table.MaxN and Table.MinN return stable ordered row windows by one column',()=>{
 const input=[['Team','Score'],['B',4],['A',null],['C',4],['D',2]];
 assert.deepEqual(executePowerQueryM(input,'let Result = Table.MaxN(Source, "Score", 2) in Result'),[['Team','Score'],['B',4],['C',4]]);
 assert.deepEqual(executePowerQueryM(input,'let Result = Table.MinN(Source, "Score", 2) in Result'),[['Team','Score'],['A',null],['D',2]]);
 assert.deepEqual(executePowerQueryM(input,'let Result = Table.MaxN(Source, "Score", 0) in Result'),[['Team','Score']]);
 assert.deepEqual(executePowerQueryM(input,'let Result = Table.MaxN(Source, "Score", each [Score] >= 4) in Result'),[['Team','Score'],['B',4],['C',4]]);
 assert.throws(()=>executePowerQueryM(input,'let Result = Table.MaxN(Source, "Missing", 1) in Result'),/existing column name/);
 assert.throws(()=>executePowerQueryM(input,'let Result = Table.MinN(Source, "Score", each [Score]) in Result'),/condition must return true or false/);
 assert.throws(()=>executePowerQueryM([['Value'],[1],['text']],'let Result = Table.MinN(Source, "Value", 1) in Result'),/one comparable scalar type/);
});
test('bounded M Table.RemoveMatchingRows removes typed complete-row matches',()=>{
 const input=[['ID','Name'],[1,'Ada'],[2,'Lin'],[1,'Ada'],['1','Ada']];
 assert.deepEqual(executePowerQueryM(input,'let Result = Table.RemoveMatchingRows(Source, {{1,"Ada"}, [ID=2,Name="Lin"]}) in Result'),[['ID','Name'],['1','Ada']]);
 assert.deepEqual(executePowerQueryM(input,'let Result = Table.RemoveMatchingRows(Source, {}) in Result'),input);
 assert.throws(()=>executePowerQueryM(input,'let Result = Table.RemoveMatchingRows(Source, {{1}}) in Result'),/match the table width/);
 assert.throws(()=>executePowerQueryM(input,'let Result = Table.RemoveMatchingRows(Source, {[ID=1]}) in Result'),/exactly the source columns/);
});
test('bounded M Table.InsertRows inserts exact-width scalar lists and complete records',()=>{
 const input=[['ID','Name'],[1,'Ada'],[3,'Kim']];
 assert.deepEqual(executePowerQueryM(input,'let Result = Table.InsertRows(Source, 1, {[ID=2, Name="Lin"], {4, "Mia"}}) in Result'),[['ID','Name'],[1,'Ada'],[2,'Lin'],[4,'Mia'],[3,'Kim']]);
 assert.deepEqual(executePowerQueryM(input,'let Result = Table.InsertRows(Source, 0, {}) in Result'),input);
 assert.throws(()=>executePowerQueryM(input,'let Result = Table.InsertRows(Source, 9, {{2, "Lin"}}) in Result'),/in-range whole-number offset/);
 assert.throws(()=>executePowerQueryM(input,'let Result = Table.InsertRows(Source, 0, {{2}}) in Result'),/match the source width/);
 assert.throws(()=>executePowerQueryM(input,'let Result = Table.InsertRows(Source, 0, {[ID=2]}) in Result'),/exactly the source columns/);
});
test('bounded M Table.ReplaceRows replaces an in-range slice with validated scalar rows',()=>{
 const input=[['ID','Name'],[1,'Ada'],[2,'Lin'],[3,'Kim']];
 assert.deepEqual(executePowerQueryM(input,'let Result = Table.ReplaceRows(Source, 1, 2, {[ID=8, Name="Sam"]}) in Result'),[['ID','Name'],[1,'Ada'],[8,'Sam']]);
 assert.deepEqual(executePowerQueryM(input,'let Result = Table.ReplaceRows(Source, 1, 1, {{8, "Sam"}, {9, "Jo"}}) in Result'),[['ID','Name'],[1,'Ada'],[8,'Sam'],[9,'Jo'],[3,'Kim']]);
 assert.deepEqual(executePowerQueryM(input,'let Result = Table.ReplaceRows(Source, 3, 0, {{4, "Max"}}) in Result'),[['ID','Name'],[1,'Ada'],[2,'Lin'],[3,'Kim'],[4,'Max']]);
 assert.deepEqual(executePowerQueryM(input,'let Result = Table.ReplaceRows(Source, 1, 1, {}) in Result'),[['ID','Name'],[1,'Ada'],[3,'Kim']]);
 assert.throws(()=>executePowerQueryM(input,'let Result = Table.ReplaceRows(Source, 3, 1, {{4, "Max"}}) in Result'),/range must fit/);
 assert.throws(()=>executePowerQueryM(input,'let Result = Table.ReplaceRows(Source, 1, -1, {}) in Result'),/non-negative count/);
 assert.throws(()=>executePowerQueryM(input,'let Result = Table.ReplaceRows(Source, 0, 1, {{4}}) in Result'),/match the source width/);
});
test('bounded M Table.ReverseRows reverses row order without changing schema',()=>{
 const input=[['ID','Name'],[1,'Ada'],[2,'Lin'],[3,'Kim']];
 assert.deepEqual(executePowerQueryM(input,'let Result = Table.ReverseRows(Source) in Result'),[['ID','Name'],[3,'Kim'],[2,'Lin'],[1,'Ada']]);
 assert.deepEqual(executePowerQueryM([['ID']], 'let Result = Table.ReverseRows(Source) in Result'),[['ID']]);
 assert.throws(()=>executePowerQueryM(input,'let Result = Table.ReverseRows(Source, 2) in Result'),/expects 1 arguments/);
});
test('bounded M Table.TransformRows returns safe scalar or flat-record projections',()=>{
 const result=executePowerQueryM([['ID','Name'],[1,'Ada'],[2,'Lin']], 'let Rows = Table.TransformRows(Source, each [Key = [ID], Label = [Name] & "-row"]), Result = Table.FromRecords(Rows) in Result');
 assert.deepEqual(result,[['Key','Label'],[1,'Ada-row'],[2,'Lin-row']]);
 assert.throws(()=>executePowerQueryM([['ID'],[1]], 'let Rows = Table.TransformRows(Source, each [Nested = [Value = [ID]]]), Result = Table.FromRecords(Rows) in Result'),/record of scalar values/);
 assert.throws(()=>executePowerQueryM([['ID'],[1]], 'let Rows = Table.TransformRows(Source, 1) in #table({"Count"}, {{List.Count(Rows)}})'),/each row transformation/);
});
test('bounded M applies optional types on added and transformed columns',()=>{
 const added=executePowerQueryM([['Amount'],['4']], 'let Added = Table.AddColumn(Source, "Typed", each [Amount], type number) in Added');
 const transformed=executePowerQueryM([['Amount'],['4']], 'let Converted = Table.TransformColumns(Source, {{"Amount", Number.From, type number}}) in Converted');
 assert.deepEqual(added,[['Amount','Typed'],['4',4]]);assert.deepEqual(transformed,[['Amount'],[4]]);
});
test('bounded M index columns honor offsets, increments, type checks, and unique names',()=>{
 const input=[['Name'],['Ada'],['Lin']];
 assert.deepEqual(executePowerQueryM(input,'let Indexed = Table.AddIndexColumn(Source, "RowNo", 10, 5, type number) in Indexed'),[['Name','RowNo'],['Ada',10],['Lin',15]]);
 assert.deepEqual(executePowerQueryM(input,'let Indexed = Table.AddIndexColumn(Source, "RowNo", -1, -2) in Indexed'),[['Name','RowNo'],['Ada',-1],['Lin',-3]]);
 assert.throws(()=>executePowerQueryM(input,'let Indexed = Table.AddIndexColumn(Source, "name") in Indexed'),/unique name/);
 assert.throws(()=>executePowerQueryM(input,'let Indexed = Table.AddIndexColumn(Source, "RowNo", 1, 1, type text) in Indexed'),/only number or any/);
 assert.throws(()=>executePowerQueryM(input,'let Indexed = Table.AddIndexColumn(Source, "RowNo", 1e308, 1e308) in Indexed'),/non-finite index/);
 assert.throws(()=>executePowerQueryM(input,'let Added = Table.AddColumn(Source, "name", each [Name]) in Added'),/new name/);
});
test('bounded M Table.Group applies optional scalar output types and rejects nested output types',()=>{
 const input=[['Team','Amount'],['A',4],['A',5],['B',3]];
 assert.deepEqual(executePowerQueryM(input,'let Grouped = Table.Group(Source, {"Team"}, {{"Total", each List.Sum([Amount]), type text}}) in Grouped'),[['Team','Total'],['A','9'],['B','3']]);
 assert.throws(()=>executePowerQueryM(input,'let Grouped = Table.Group(Source, {"Team"}, {{"Total", each List.Sum([Amount]), type table}}) in Grouped'),/scalar output type/);
 assert.throws(()=>executePowerQueryM(input,'let Grouped = Table.Group(Source, {"Team"}, {{"Total", each List.Sum([Amount]), type logical}}) in Grouped'),/convert value/);
 assert.throws(()=>executePowerQueryM(input,'let Grouped = Table.Group(Source, {"Team"}, {{"Total", each List.Sum([Amount])}}, 1) in Grouped'),/comparer and group-kind options are not supported/);
});
test('bounded M supports explicit first/last and remove row slices',()=>{
 const input=[['N'],[1],[2],[3],[4]],result=executePowerQueryM(input,'let Last = Table.LastN(Source, 2), First = Table.RemoveLastN(Last, 1) in First');
 assert.deepEqual(result,[['N'],[3]]);
 assert.deepEqual(executePowerQueryM(input,'let Kept = Table.RemoveFirstN(Source, 2) in Kept'),[['N'],[3],[4]]);
 assert.deepEqual(executePowerQueryM(input,'let Last = Table.LastN(Source, 0) in Last'),[['N']]);
 assert.deepEqual(executePowerQueryM(input,'let First = Table.FirstN(Source, each [N] <= 2) in First'),[['N'],[1],[2]]);
 assert.deepEqual(executePowerQueryM(input,'let Last = Table.LastN(Source, each [N] > 2) in Last'),[['N'],[3],[4]]);
 assert.deepEqual(executePowerQueryM(input,'let Skipped = Table.Skip(Source, each [N] < 3) in Skipped'),[['N'],[3],[4]]);
 assert.deepEqual(executePowerQueryM(input,'let Removed = Table.RemoveFirstN(Source, each [N] < 3) in Removed'),[['N'],[3],[4]]);
 assert.deepEqual(executePowerQueryM(input,'let Removed = Table.RemoveLastN(Source, each [N] > 2) in Removed'),[['N'],[1],[2]]);
 assert.deepEqual(executePowerQueryM(input,'let Skipped = Table.Skip(Source) in Skipped'),[['N'],[2],[3],[4]]);
 assert.deepEqual(executePowerQueryM(input,'let Removed = Table.RemoveFirstN(Source) in Removed'),[['N'],[2],[3],[4]]);
 assert.deepEqual(executePowerQueryM(input,'let Removed = Table.RemoveLastN(Source) in Removed'),[['N'],[1],[2],[3]]);
 assert.throws(()=>executePowerQueryM(input,'let Bad = Table.FirstN(Source, each [N]) in Bad'),/condition must return true or false/);
});
test('bounded M Table.Range returns remaining rows when its nullable count is omitted',()=>{
 const input=[['N'],[1],[2],[3]];
 assert.deepEqual(executePowerQueryM(input,'let Result = Table.Range(Source, 1) in Result'),[['N'],[2],[3]]);
 assert.deepEqual(executePowerQueryM(input,'let Result = Table.Range(Source, 1, null) in Result'),[['N'],[2],[3]]);
 assert.deepEqual(executePowerQueryM(input,'let Result = Table.Range(Source, 1, 0) in Result'),[['N']]);
 assert.throws(()=>executePowerQueryM(input,'let Result = Table.Range(Source, 1, 1.5) in Result'),/whole-number offset and count/);
});
test('bounded M Table.AlternateRows keeps the prefix then alternates skipped and taken rows',()=>{
 const input=[['N'],[1],[2],[3],[4],[5],[6]];
 assert.deepEqual(executePowerQueryM(input,'let Result = Table.AlternateRows(Source, 1, 1, 1) in Result'),[['N'],[1],[3],[5]]);
 assert.deepEqual(executePowerQueryM(input,'let Result = Table.AlternateRows(Source, 2, 0, 2) in Result'),[['N'],[1],[2],[3],[4],[5],[6]]);
 assert.deepEqual(executePowerQueryM(input,'let Result = Table.AlternateRows(Source, 2, 1, 0) in Result'),[['N'],[1],[2]]);
 assert.throws(()=>executePowerQueryM(input,'let Result = Table.AlternateRows(Source, 0, -1, 1) in Result'),/non-negative whole-number/);
});
test('bounded M Table.Repeat repeats ordered rows within row and cell limits',()=>{
 const input=[['N','Label'],[1,'a'],[2,'b']];
 assert.deepEqual(executePowerQueryM(input,'let Result = Table.Repeat(Source, 2) in Result'),[['N','Label'],[1,'a'],[2,'b'],[1,'a'],[2,'b']]);
 assert.deepEqual(executePowerQueryM(input,'let Result = Table.Repeat(Source, 0) in Result'),[['N','Label']]);
 assert.throws(()=>executePowerQueryM(input,'let Result = Table.Repeat(Source, 1.5) in Result'),/whole number/);
 assert.throws(()=>executePowerQueryM(input,'let Result = Table.Repeat(Source, 501) in Result'),/row or cell limit/);
});
test('bounded M Table.SplitAt returns two schema-preserving table values for list consumers',()=>{
 const input=[['N'],[1],[2],[3]];
 assert.deepEqual(executePowerQueryM(input,'let Parts = Table.SplitAt(Source, 2), Head = List.First(Parts), Tail = List.Last(Parts) in #table({"HeadRows","TailRows","Parts"}, {{Table.RowCount(Head),Table.RowCount(Tail),List.Count(Parts)}})'),[['HeadRows','TailRows','Parts'],[2,1,2]]);
 assert.deepEqual(executePowerQueryM(input,'let Parts = Table.SplitAt(Source, 0), Head = List.First(Parts), Tail = List.Last(Parts) in #table({"HeadRows","TailRows"}, {{Table.RowCount(Head),Table.RowCount(Tail)}})'),[['HeadRows','TailRows'],[0,3]]);
 assert.throws(()=>executePowerQueryM(input,'let Parts = Table.SplitAt(Source, -1), Count = List.Count(Parts) in #table({"Count"}, {{Count}})'),/whole number/);
});
test('bounded M Table.Split creates ordered schema-preserving pages for list consumers',()=>{
 const input=[['N'],[1],[2],[3],[4],[5]];
 assert.deepEqual(executePowerQueryM(input,'let Pages = Table.Split(Source, 2), First = List.First(Pages), Last = List.Last(Pages) in #table({"PageCount","FirstRows","LastRows"}, {{List.Count(Pages),Table.RowCount(First),Table.RowCount(Last)}})'),[['PageCount','FirstRows','LastRows'],[3,2,1]]);
 assert.deepEqual(executePowerQueryM([['N']],'let Pages = Table.Split(Source, 2) in #table({"PageCount"}, {{List.Count(Pages)}})'),[['PageCount'],[0]]);
 assert.throws(()=>executePowerQueryM(input,'let Pages = Table.Split(Source, 0), Count = List.Count(Pages) in #table({"Count"}, {{Count}})'),/page size/);
});
test('bounded M Table.ToList joins text rows with the supported delimiter combiner',()=>{
 const input=[['A','B'],['hello','world'],['a,b','c"d']];
 assert.deepEqual(executePowerQueryM(input,'let Values = Table.ToList(Source, Combiner.CombineTextByDelimiter("-")), Result = #table({"Count","First","Last"}, {{List.Count(Values),List.First(Values),List.Last(Values)}}) in Result'),[['Count','First','Last'],[2,'hello-world','a,b-c"d']]);
 assert.deepEqual(executePowerQueryM(input,'let Values = Table.ToList(Source, Combiner.CombineTextByDelimiter(",", QuoteStyle.Csv)), Result = #table({"First","Last"}, {{List.First(Values),List.Last(Values)}}) in Result'),[['First','Last'],['hello,world','"a,b","c""d"']]);
 assert.deepEqual(executePowerQueryM([['A','B']],'let Values = Table.ToList(Source) in #table({"Count"}, {{List.Count(Values)}})'),[['Count'],[0]]);
 assert.throws(()=>executePowerQueryM([['A'],[1]],'let Values = Table.ToList(Source), Count = List.Count(Values) in #table({"Count"}, {{Count}})'),/must contain text/);
});
test('bounded M Table.FromList supports comma splitting and the delimiter/no-split generators',()=>{
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let Result = Table.FromList({"a,apple", "b,ball"}, null, {"Letter","Word"}) in Result'),[['Letter','Word'],['a','apple'],['b','ball']]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let Result = Table.FromList({"a-b", "c-d"}, Splitter.SplitTextByDelimiter("-"), {"Left","Right"}) in Result'),[['Left','Right'],['a','b'],['c','d']]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let Result = Table.FromList({1, 2}, Splitter.SplitByNothing(), {"Value"}) in Result'),[['Value'],[1],[2]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Result = Table.FromList({"a,b,c"}, null, {"A","B"}) in Result'),/only 2 columns were provided/);
});
test('bounded M converts records to ordered value lists and splits record lists into tables',()=>{
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let Values = Record.ToList([A=1,B="x"]), Result = #table({"Count","First","Last"}, {{List.Count(Values),List.First(Values),List.Last(Values)}}) in Result'),[['Count','First','Last'],[2,1,'x']]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let Values = Record.FieldNames(["2"="second","1"="first"]), Result = #table({"First","Last"}, {{List.First(Values),List.Last(Values)}}) in Result'),[['First','Last'],['2','1']]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let Values = Record.FieldValues([A=null,B=2]), Result = #table({"Count","First","Last"}, {{List.Count(Values),List.First(Values),List.Last(Values)}}) in Result'),[['Count','First','Last'],[2,null,2]]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let Result = Table.FromList({[A=1,B="x"],[A=2,B="y"]}, Record.FieldValues, {"ID","Name"}) in Result'),[['ID','Name'],[1,'x'],[2,'y']]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Result = Record.FieldValues({1,2}) in Result'),/expects a record/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Result = Record.ToList(Source) in Result'),/expects a record/);
});
test('bounded M record helpers expose names, exact field access, presence and defaults',()=>{
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let R = [Name="Ada", Score=7], Names = Record.FieldNames(R), Result = #table({"Name1","Name2","Field","HasAll","HasNone","Default","NullDefault"}, {{List.First(Names),List.Last(Names),Record.Field(R,"Score"),Record.HasFields(R,{"Name","Score"}),Record.HasFields(R,{"Name","Missing"}),Record.FieldOrDefault(R,"Missing",42),Record.FieldOrDefault(R,"Missing")}}) in Result'),[['Name1','Name2','Field','HasAll','HasNone','Default','NullDefault'],['Name','Score',7,true,false,42,null]]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let Names = Record.FieldNames([A=1,a=2]), Result = #table({"Count","First","Last"}, {{Record.FieldCount([A=1,a=2]),List.First(Names),List.Last(Names)}}) in Result'),[['Count','First','Last'],[2,'A','a']]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let R = [headers="H",rows="R"], Result = #table({"Count","Rows"}, {{Record.FieldCount(R),Record.Field(R,"rows")}}) in Result'),[['Count','Rows'],[2,'R']]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let R = Record.FromList({{"A"},{1}},{"headers","rows"}), Result = #table({"Count","HeaderValue"}, {{Record.FieldCount(R),List.First(Record.Field(R,"headers"))}}) in Result'),[['Count','HeaderValue'],[2,'A']]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let R = Record.FromList({7},{"__proto__"}), Result = #table({"Count","Value"}, {{Record.FieldCount(R),Record.Field(R,"__proto__")}}) in Result'),[['Count','Value'],[1,7]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let R = [A=1,A=2] in R'),/field names must be unique/);
 const tooManyFields=Array.from({length:105},(_,index)=>`F${index}=${index}`).join(',');
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],`let R = [${tooManyFields}] in R`),/records are limited to 104 fields/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Result = Record.Field([Name="Ada"], "name") in Result'),/cannot find field/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Result = Record.FieldOrDefault([A=1], 2) in Result'),/field name must be text/);
});
test('bounded M Record.SelectFields and RemoveFields preserve order and validate missing-field policies',()=>{
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let R = [A=1,B=2,C=3], Selected = Record.SelectFields(R,{"C","A","Missing"},MissingField.UseNull), Removed = Record.RemoveFields(R,{"B","Missing"},MissingField.Ignore), SN = Record.FieldNames(Selected), RN = Record.FieldNames(Removed), Result = #table({"SelectFirst","SelectSecond","SelectLast","SelectMissing","RemoveFirst","RemoveLast"}, {{List.First(SN),List.First(List.Skip(SN,1)),List.Last(SN),Record.Field(Selected,"Missing"),List.First(RN),List.Last(RN)}}) in Result'),[['SelectFirst','SelectSecond','SelectLast','SelectMissing','RemoveFirst','RemoveLast'],['A','C','Missing',null,'A','C']]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let R = Record.SelectFields([A=1],"Missing") in R'),/cannot find field/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let R = Record.RemoveFields([A=1],"Missing") in R'),/cannot find field/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let R = Record.RemoveFields([A=1],"A",MissingField.UseNull) in R'),/supports MissingField.Error or MissingField.Ignore/);
});
test('bounded M record mutation, ordered rename/reorder, and right-biased combine operations',()=>{
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let R = Record.AddField([A=1],"B",2), Swapped = Record.RenameFields(R,{{"A","B"},{"B","A"}}), Names = Record.FieldNames(Swapped), Reordered = Record.ReorderFields(Swapped,{"A","Missing"},MissingField.UseNull), OutNames = Record.FieldNames(Reordered), Result = #table({"Name1","Name2","A","Missing","B"}, {{List.First(Names),List.Last(Names),Record.Field(Reordered,"A"),Record.Field(Reordered,"Missing"),Record.Field(Reordered,"B")}}) in Result'),[['Name1','Name2','A','Missing','B'],['B','A',2,null,1]]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let R = Record.Combine({[A=1,B=2],[A=9,C=3]}), Result = #table({"Count","A","B","C","Names"}, {{Record.FieldCount(R),Record.Field(R,"A"),Record.Field(R,"B"),Record.Field(R,"C"),List.Last(Record.FieldNames(R))}}) in Result'),[['Count','A','B','C','Names'],[3,9,2,3,'C']]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let R = Record.RenameFields([A=1],{"Missing","New"},MissingField.UseNull), Result = #table({"Count","Old","New"}, {{Record.FieldCount(R),Record.Field(R,"A"),Record.Field(R,"New")}}) in Result'),[['Count','Old','New'],[2,1,null]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let R = Record.AddField([A=1],"A",2) in R'),/duplicate field/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let R = Record.RenameFields([A=1,B=2],{{"A","B"}}) in R'),/must be unique/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let R = Record.AddField([A=1],"B",2,true) in R'),/delayed evaluation is not supported/);
});
test('bounded M structural equality and overloaded list/record concatenation follow M value kinds',()=>{
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let RecordEqual = [A=1,B={2,3}] = [B={2,3},A=1], ListEqual = {1,[A=2]} = {1,[A=2]}, ListDifferent = {1,2} <> {2,1}, Merged = [A=1,B=2] & [A=9,C=3], Joined = {1,2} & {3}, Result = #table({"RecordEqual","ListEqual","ListDifferent","A","B","C","JoinedLast"}, {{RecordEqual,ListEqual,ListDifferent,Record.Field(Merged,"A"),Record.Field(Merged,"B"),Record.Field(Merged,"C"),List.Last(Joined)}}) in Result'),[['RecordEqual','ListEqual','ListDifferent','A','B','C','JoinedLast'],[true,true,true,9,2,3,3]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Result = 1 & "one" in #table({"Value"}, {{Result}})'),/two text, list, or record values/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let Values = List.Repeat({1},12000) & {2} in #table({"Count"}, {{List.Count(Values)}})'),/list concatenation exceeds the 12,000-item limit/);
});
test('bounded M Record.TransformFields applies functions and each lambdas with checked missing fields',()=>{
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let R = Record.TransformFields([Name=" ada ",Amount="42"],{{"Name",Text.Upper},{"Amount",each Number.From(_) + 1}}), Result = #table({"Name","Amount"}, {{Record.Field(R,"Name"),Record.Field(R,"Amount")}}) in Result'),[['Name','Amount'],[' ADA ',43]]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let R = Record.TransformFields([A="x"],{{"Missing",Text.Upper}},MissingField.UseNull) in #table({"Fields","Value"}, {{Record.FieldCount(R),Record.Field(R,"Missing")}})'),[['Fields','Value'],[2,null]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let R = Record.TransformFields([A="x"],{{"Missing",Text.Upper}}) in R'),/cannot find field/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let R = Record.TransformFields([A="x"],{{"A",each {1}}}) in R'),/finite scalar or null/);
});
test('bounded M Record.FieldCount and Record.ToTable preserve the record schema and order',()=>{
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let R = [Key="A",Value=7], T = Record.ToTable(R), Names = Table.Column(T,"Name"), Values = Table.Column(T,"Value"), EmptyCount = Record.FieldCount([]), Result = #table({"Count","FirstName","LastName","FirstValue","LastValue","Empty"}, {{Record.FieldCount(R),List.First(Names),List.Last(Names),List.First(Values),List.Last(Values),EmptyCount}}) in Result'),[['Count','FirstName','LastName','FirstValue','LastValue','Empty'],[2,'Key','Value','A',7,0]]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let T = Record.ToTable([Value=#duration(0,1,0,0)]) in T'),/scalar field values/);
});
test('bounded M Record.FromList and Record.FromTable reconstruct ordered scalar records',()=>{
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let R = Record.FromList({7,"Ada"},{"ID","Name"}), T = Record.ToTable(R), Again = Record.FromTable(T), Names = Record.FieldNames(Again), Result = #table({"NameFirst","NameLast","Count","ID","Value"}, {{List.First(Names),List.Last(Names),Record.FieldCount(Again),Record.Field(Again,"ID"),Record.Field(Again,"Name")}}) in Result'),[['NameFirst','NameLast','Count','ID','Value'],['ID','Name',2,7,'Ada']]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let R = Record.FromTable(#table({"Value","Name"}, {{null,"A"},{false,"B"}})), Values = Record.FieldValues(R), Result = #table({"Count","First","Last"}, {{Record.FieldCount(R),List.First(Values),List.Last(Values)}}) in Result'),[['Count','First','Last'],[2,null,false]]);
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let T = Table.FromRecords({["2"="two","1"="one"]}), Names = Table.ColumnNames(T), Roundtrip = List.First(Table.ToRecords(T)), Fields = Record.FieldNames(Roundtrip), Result = #table({"ColumnFirst","ColumnLast","RecordFirst","RecordLast"}, {{List.First(Names),List.Last(Names),List.First(Fields),List.Last(Fields)}}) in Result'),[['ColumnFirst','ColumnLast','RecordFirst','RecordLast'],['2','1','2','1']]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let R = Record.FromList({1},{"A","B"}) in R'),/equally sized lists/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let R = Record.FromList({1,2},{"A","A"}) in R'),/unique/);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],'let R = Record.FromTable(#table({"Name","Value"}, {{"A",1},{"A",2}})) in R'),/occurs more than once/);
});
test('bounded M provides Unicode-safe text slicing, padding, reversal and scalar conversion',()=>{
 const result=executePowerQueryM([['Word'],['😀abc']], 'let Start = Table.AddColumn(Source, "Start", each Text.Start([Word], 2)), End = Table.AddColumn(Start, "End", each Text.End([Word], 2)), Reversed = Table.AddColumn(End, "Reverse", each Text.Reverse([Word])), Padded = Table.AddColumn(Reversed, "Padded", each Text.PadStart([Word], 6, "😀")), Number = Table.AddColumn(Padded, "NumberText", each Text.From(42)) in Number');
 assert.deepEqual(result[1],['😀abc','😀a','bc','cba😀','😀😀😀abc','42']);
 assert.deepEqual(executePowerQueryM([['Text'],['A\u0085B']], 'let Clean = Table.AddColumn(Source, "Clean", each Text.Clean([Text])) in Clean'),[['Text','Clean'],['A\u0085B','AB']]);
 assert.deepEqual(executePowerQueryM([['Text'],['...value...']], 'let Trimmed = Table.AddColumn(Source, "Trim", each Text.Trim([Text], ".")), Start = Table.AddColumn(Trimmed, "Start", each Text.TrimStart([Text], {"."})), End = Table.AddColumn(Start, "End", each Text.TrimEnd([Text], ".")) in End'),[['Text','Trim','Start','End'],['...value...','value','value...','...value']]);
 assert.deepEqual(executePowerQueryM([['Text'],['A😀BC']], 'let Slice = Table.AddColumn(Source, "Slice", each Text.Range([Text], 1, 2)), Tail = Table.AddColumn(Slice, "Tail", each Text.Range([Text], 3)) in Tail'),[['Text','Slice','Tail'],['A😀BC','😀B','C']]);
 assert.throws(()=>executePowerQueryM([['Text'],['abc']], 'let Slice = Table.AddColumn(Source, "Slice", each Text.Range([Text], -1)) in Slice'),/non-negative/);
});
test('bounded M Text.ToList returns Unicode code points under the list-size cap',()=>{
 assert.deepEqual(executePowerQueryM([['Seed'],[1]],'let Characters = Text.ToList("A😀é"), Result = #table({"Count","First","Second","Last"}, {{List.Count(Characters),List.First(Characters),List.First(List.Skip(Characters, 1)),List.Last(Characters)}}) in Result'),[['Count','First','Second','Last'],[4,'A','😀','́']]);
 assert.throws(()=>executePowerQueryM([['Seed'],[1]],`let Characters = Text.ToList("${'a'.repeat(12001)}"), Count = List.Count(Characters) in #table({"Count"}, {{Count}})`),/limited to 12000 characters/);
});
test('bounded M extracts text around the first delimiter and errors clearly when absent',()=>{
 const result=executePowerQueryM([['Path'],['north|central|south']], 'let Before = Table.AddColumn(Source, "Before", each Text.BeforeDelimiter([Path], "|")), After = Table.AddColumn(Before, "After", each Text.AfterDelimiter([Path], "|")) in After');
 assert.deepEqual(result,[['Path','Before','After'],['north|central|south','north','central|south']]);
 assert.throws(()=>executePowerQueryM([['Path'],['no marker']], 'let Before = Table.AddColumn(Source, "Before", each Text.BeforeDelimiter([Path], "|")) in Before'),/could not find/);
 assert.deepEqual(executePowerQueryM([['Path'],['start:value:end']], 'let Between = Table.AddColumn(Source, "Value", each Text.BetweenDelimiters([Path], "start:", ":end")) in Between'),[['Path','Value'],['start:value:end','value']]);
 assert.throws(()=>executePowerQueryM([['Path'],['start:value']], 'let Between = Table.AddColumn(Source, "Value", each Text.BetweenDelimiters([Path], "start:", ":end")) in Between'),/end delimiter/);
 assert.deepEqual(executePowerQueryM([['Text'],['a-b-c']], 'let First = Table.AddColumn(Source, "First", each Text.PositionOf([Text], "-")), Last = Table.AddColumn(First, "Last", each Text.PositionOf([Text], "-", Occurrence.Last)) in Last'),[['Text','First','Last'],['a-b-c',1,3]]);
 assert.deepEqual(executePowerQueryM([['Text'],['abc']], 'let Position = Table.AddColumn(Source, "Position", each Text.PositionOf([Text], "z")) in Position'),[['Text','Position'],['abc',-1]]);
 assert.deepEqual(executePowerQueryM([['Text'],['a+b+c']], 'let Before = Table.AddColumn(Source, "Before", each Text.BeforeDelimiter([Text], "+")), After = Table.AddColumn(Before, "After", each Text.AfterDelimiter([Text], "+")) in After'),[['Text','Before','After'],['a+b+c','a','b+c']]);
 assert.deepEqual(executePowerQueryM([['Text'],['north-south-east']], 'let Before = Table.AddColumn(Source, "Before", each Text.BeforeDelimiter([Text], "-")), After = Table.AddColumn(Before, "After", each Text.AfterDelimiter([Text], "-")) in After'),[['Text','Before','After'],['north-south-east','north','south-east']]);
 assert.deepEqual(executePowerQueryM([['Text'],['a1😀b2']], 'let Selected = Table.AddColumn(Source, "Keep", each Text.Select([Text], {"a", "😀", "2"})), Removed = Table.AddColumn(Selected, "Drop", each Text.Remove([Text], "12")) in Removed'),[['Text','Keep','Drop'],['a1😀b2','a😀2','a😀b']]);
});
test('saved offline query replays cleanup and conversion steps into a separate load range',()=>{
 const book=sample(),spec=validateQuerySpec({id:'query-customer-cleanup',name:'Customer cleanup',sourceRange:'A1:B3',destination:'D1',steps:['trim','upper','double']},book.rows,book.cols),result=buildQueryRefresh(book,spec);
 assert.equal(result.outputRange,'D1:E3');assert.deepEqual(['D1','E1','D2','E2','D3','E3'].map(ref=>book.getRaw(ref)),['','','','','','']);book.applyEdits(result.edits,false);assert.deepEqual(['D1','E1','D2','E2','D3','E3'].map(ref=>book.getValue(ref).value),['NAME','AMOUNT','ALICE',24,'BOB',14]);
});
test('query refresh picks up changed source values without changing the source during preview',()=>{
 const book=sample(),spec={id:'query-clean',name:'Clean',sourceRange:'A1:B3',destination:'D1',steps:['trim','upper'],lastOutputRange:'D1:E3'};book.applyEdits(buildQueryRefresh(book,spec).edits,false);book.applyEdits([{ref:'A2',raw:'  eve '}],false);const result=buildQueryRefresh(book,spec);assert.equal(book.getValue('A2').value,'  eve ');book.applyEdits(result.edits,false);assert.equal(book.getValue('D2').value,'EVE');
});
test('query definitions persist in GridForge workbook JSON and reject unsafe steps or overlap',()=>{
 const book=sample(),doc=new SpreadsheetDocument(book),query={id:'query-clean',name:'Clean',sourceRange:'A1:B3',destination:'D1',steps:['trim']};doc.active.tools.queries=validateTools({queries:[query]},book.rows,book.cols).queries;const restored=SpreadsheetDocument.fromJSON(doc.toJSON());assert.equal(restored.active.tools.queries[0]?.name,'Clean');assert.throws(()=>validateQuerySpec({...query,steps:['insert-sum']},book.rows,book.cols),/supported/);assert.throws(()=>validateQuerySpec({...query,destination:'B1'},book.rows,book.cols),/overlap/);
});
test('XLSX export refreshes saved query results without changing the open workbook',async()=>{
 const book=sample(),doc=new SpreadsheetDocument(book),spec={id:'query-clean',name:'Clean',sourceRange:'A1:B3',destination:'D1',steps:['trim','upper'],lastOutputRange:'D1:E3'};doc.active.tools.queries=validateTools({queries:[spec]},book.rows,book.cols).queries;book.applyEdits([{ref:'A2',raw:'  eve '}],false);
 const parts=await zipRead(exportXlsx(doc)),xml=parts.get('xl/worksheets/sheet1.xml');assert.match(xml,/<c r="D2" t="inlineStr"><is><t xml:space="preserve">EVE<\/t><\/is><\/c>/);assert.equal(book.getValue('D2').value,null);assert.equal(book.getValue('A2').value,'  eve ');
});

