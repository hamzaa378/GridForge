/** v10.4: registered Excel-connected function names.
 * These remaining functions require data sources GridForge intentionally does not simulate.
 * GETPIVOTDATA is handled separately by the workbook-aware formula evaluator.
 * DO NOT claim service-dependent functions are executable offline Excel equivalents.
 */
import {SheetError} from './functions.js';
import type {FunctionDefinition,Datum} from './functions.js';
export const connectedFunctionRequirements: Record<string,string> = {
  CUBEKPIMEMBER: "OLAP cube connection",
  CUBEMEMBER: "OLAP cube connection",
  CUBEMEMBERPROPERTY: "OLAP cube connection",
  CUBERANKEDMEMBER: "OLAP cube connection",
  CUBESET: "OLAP cube connection",
  CUBEVALUE: "OLAP cube connection",
  RTD: "Windows RTD COM server",
  STOCKHISTORY: "licensed market-data provider",
  WEBSERVICE: "authorized network service",
};
export const connectedFunctionNames = Object.keys(connectedFunctionRequirements).sort();
export const connectedFunctionDefinitions: Record<string,FunctionDefinition> = {
  CUBEKPIMEMBER: {description:"Connected Excel function \u2014 requires OLAP cube connection. Recognized for workbook compatibility but not available in GridForge's offline formula engine.",example:"=CUBEKPIMEMBER(\"Sales\",\"Revenue\",1)",run:(args:Datum[])=>{
    if(args.length<3||args.length>3) throw new SheetError('#VALUE!',"CUBEKPIMEMBER expects exactly 3 arguments.");
    throw new SheetError('#NAME?',"CUBEKPIMEMBER requires OLAP cube connection; use an appropriately connected copy of Microsoft Excel to evaluate it.");
  }},
  CUBEMEMBER: {description:"Connected Excel function \u2014 requires OLAP cube connection. Recognized for workbook compatibility but not available in GridForge's offline formula engine.",example:"=CUBEMEMBER(\"Sales\",\"[Region].[North]\")",run:(args:Datum[])=>{
    if(args.length<2||args.length>3) throw new SheetError('#VALUE!',"CUBEMEMBER expects 2\u20133 arguments.");
    throw new SheetError('#NAME?',"CUBEMEMBER requires OLAP cube connection; use an appropriately connected copy of Microsoft Excel to evaluate it.");
  }},
  CUBEMEMBERPROPERTY: {description:"Connected Excel function \u2014 requires OLAP cube connection. Recognized for workbook compatibility but not available in GridForge's offline formula engine.",example:"=CUBEMEMBERPROPERTY(\"Sales\",\"[Region].[North]\",\"Caption\")",run:(args:Datum[])=>{
    if(args.length<3||args.length>3) throw new SheetError('#VALUE!',"CUBEMEMBERPROPERTY expects exactly 3 arguments.");
    throw new SheetError('#NAME?',"CUBEMEMBERPROPERTY requires OLAP cube connection; use an appropriately connected copy of Microsoft Excel to evaluate it.");
  }},
  CUBERANKEDMEMBER: {description:"Connected Excel function \u2014 requires OLAP cube connection. Recognized for workbook compatibility but not available in GridForge's offline formula engine.",example:"=CUBERANKEDMEMBER(\"Sales\",\"[Products]\",1)",run:(args:Datum[])=>{
    if(args.length<3||args.length>4) throw new SheetError('#VALUE!',"CUBERANKEDMEMBER expects 3\u20134 arguments.");
    throw new SheetError('#NAME?',"CUBERANKEDMEMBER requires OLAP cube connection; use an appropriately connected copy of Microsoft Excel to evaluate it.");
  }},
  CUBESET: {description:"Connected Excel function \u2014 requires OLAP cube connection. Recognized for workbook compatibility but not available in GridForge's offline formula engine.",example:"=CUBESET(\"Sales\",\"[Products].Members\",\"Products\")",run:(args:Datum[])=>{
    if(args.length<2||args.length>5) throw new SheetError('#VALUE!',"CUBESET expects 2\u20135 arguments.");
    throw new SheetError('#NAME?',"CUBESET requires OLAP cube connection; use an appropriately connected copy of Microsoft Excel to evaluate it.");
  }},
  CUBEVALUE: {description:"Connected Excel function \u2014 requires OLAP cube connection. Recognized for workbook compatibility but not available in GridForge's offline formula engine.",example:"=CUBEVALUE(\"Sales\",\"[Measures].[Revenue]\")",run:(args:Datum[])=>{
    if(args.length<1||args.length>99) throw new SheetError('#VALUE!',"CUBEVALUE expects 1 or more arguments.");
    throw new SheetError('#NAME?',"CUBEVALUE requires OLAP cube connection; use an appropriately connected copy of Microsoft Excel to evaluate it.");
  }},
  GETPIVOTDATA: {description:"Returns an aggregate from a saved GridForge Pivot Summary, with optional field/item criteria. Supports active report, item, label, date, and Top/Bottom value filters, configured base aggregations, and the Pivot Summary's configured Show Values As calculations.",example:"=GETPIVOTDATA(\"Sum of Sales\",A1,\"Region\",\"North\")",run:(args:Datum[])=>{
    if(args.length<2||args.length>255||(args.length-2)%2!==0) throw new SheetError('#VALUE!',"GETPIVOTDATA needs a data field, a PivotTable reference, and complete field/item pairs.");
    throw new SheetError('#REF!',"GETPIVOTDATA needs a saved generated PivotTable in its containing workbook.");
  }},
  RTD: {description:"Connected Excel function \u2014 requires Windows RTD COM server. Recognized for workbook compatibility but not available in GridForge's offline formula engine.",example:"=RTD(\"GridForge.NotInstalled\",\"\",\"quote\")",run:(args:Datum[])=>{
    if(args.length<3||args.length>99) throw new SheetError('#VALUE!',"RTD expects 3 or more arguments.");
    throw new SheetError('#N/A',"RTD requires Windows RTD COM server; use an appropriately connected copy of Microsoft Excel to evaluate it.");
  }},
  STOCKHISTORY: {description:"Connected Excel function \u2014 requires licensed market-data provider. Recognized for workbook compatibility but not available in GridForge's offline formula engine.",example:"=STOCKHISTORY(\"MSFT\",DATE(2025,1,1))",run:(args:Datum[])=>{
    if(args.length<2||args.length>11) throw new SheetError('#VALUE!',"STOCKHISTORY expects 2\u201311 arguments.");
    throw new SheetError('#CONNECT!',"STOCKHISTORY requires licensed market-data provider; use an appropriately connected copy of Microsoft Excel to evaluate it.");
  }},
  WEBSERVICE: {description:"Connected Excel function \u2014 requires authorized network service. Recognized for workbook compatibility but not available in GridForge's offline formula engine.",example:"=WEBSERVICE(\"https://example.org/\")",run:(args:Datum[])=>{
    if(args.length<1||args.length>1) throw new SheetError('#VALUE!',"WEBSERVICE expects exactly 1 argument.");
    throw new SheetError('#CONNECT!',"WEBSERVICE requires authorized network service; use an appropriately connected copy of Microsoft Excel to evaluate it.");
  }},
};
