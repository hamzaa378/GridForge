#!/usr/bin/env python3
"""Read-only real Excel COM acceptance test for v10.9 chart + table OOXML.
Requires Windows, installed Microsoft Excel and pywin32; never executes macros.
"""
import json,sys,platform,pathlib,datetime
ROOT=pathlib.Path(__file__).resolve().parent.parent
SAMPLE=ROOT/'examples'/'Excel_Migration_Demo_v10_9.xlsx'
OUT=ROOT/'docs'/'EXCEL_MIGRATION_WINDOWS_V10_9_RESULTS.json'
if sys.platform!='win32':raise SystemExit('Windows desktop with Microsoft Excel is required.')
try:import win32com.client as com
except ImportError:raise SystemExit('Install pywin32 in your build/testing environment: python -m pip install pywin32')
excel=None;book=None
try:
    excel=com.DispatchEx('Excel.Application')
    excel.Visible=False;excel.DisplayAlerts=False
    # Do not modify global calculation mode or run workbook macros.
    excel.AutomationSecurity=3 # msoAutomationSecurityForceDisable
    book=excel.Workbooks.Open(str(SAMPLE.resolve()),UpdateLinks=0,ReadOnly=True,IgnoreReadOnlyRecommended=True)
    sheets=[]
    for ws in book.Worksheets:
        chart_count=int(ws.ChartObjects().Count)
        table_count=int(ws.ListObjects.Count)
        charts=[]
        for i in range(1,chart_count+1):
            item=ws.ChartObjects(i)
            try:kind=int(item.Chart.ChartType)
            except Exception:kind=None
            charts.append({'name':str(item.Name),'excelChartType':kind})
        tables=[{'name':str(ws.ListObjects(i).Name),'style':str(ws.ListObjects(i).TableStyle)} for i in range(1,table_count+1)]
        sheets.append({'sheet':str(ws.Name),'charts':chart_count,'tables':table_count,'chartDetails':charts,'tableDetails':tables})
    total_charts=sum(x['charts'] for x in sheets);total_tables=sum(x['tables'] for x in sheets)
    result={'source':'Microsoft Excel desktop','version':str(excel.Version),'build':str(excel.Build),'workbook':SAMPLE.name,'readOnly':True,'sheets':sheets,'totalCharts':total_charts,'totalTables':total_tables,'expectedCharts':13,'expectedTables':11,'status':'PASS' if total_charts==13 and total_tables==11 else 'MISMATCH','limits':'Object counts and chart types only; not exhaustive visual/data parity.'}
    OUT.write_text(json.dumps(result,indent=2),encoding='utf-8')
    print(json.dumps({'status':result['status'],'excelVersion':result['version'],'charts':total_charts,'tables':total_tables,'output':str(OUT)},indent=2))
    sys.exit(0 if result['status']=='PASS' else 2)
finally:
    if book is not None:book.Close(SaveChanges=False)
    if excel is not None:excel.Quit()
