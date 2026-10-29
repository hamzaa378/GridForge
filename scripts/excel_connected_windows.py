#!/usr/bin/env python3
"""Windows Excel comparison for 10 known connected Excel formulas.

Default run tests only harmless missing-connection cases in a temporary blank
Excel workbook. It neither opens user workbooks nor triggers data services.
RTD / STOCKHISTORY / WEBSERVICE are intentionally skipped: their real behavior
requires user-approved external services, not a meaningful offline parity test.
"""
from __future__ import annotations
import json, os, pathlib, sys, time
ROOT=pathlib.Path(__file__).resolve().parents[1]
RESULT=ROOT/'docs'/'V10_4_CONNECTED_EXCEL_WINDOWS_RESULTS.json'
CASES=[
 ('CUBEKPIMEMBER','=CUBEKPIMEMBER("GridForgeMissing","Kpi",1)','#NAME?'),
 ('CUBEMEMBER','=CUBEMEMBER("GridForgeMissing","[Region].[North]")','#NAME?'),
 ('CUBEMEMBERPROPERTY','=CUBEMEMBERPROPERTY("GridForgeMissing","[Region].[North]","Caption")','#NAME?'),
 ('CUBERANKEDMEMBER','=CUBERANKEDMEMBER("GridForgeMissing","[Set]",1)','#NAME?'),
 ('CUBESET','=CUBESET("GridForgeMissing","[Products].Members")','#NAME?'),
 ('CUBEVALUE','=CUBEVALUE("GridForgeMissing","[Measures].[Sales]")','#NAME?'),
 ('GETPIVOTDATA','=GETPIVOTDATA("Sales",A1)','#REF!'),
 ('RTD','=RTD("GridForge.NotInstalled","","quote")','#N/A'),
 ('STOCKHISTORY','=STOCKHISTORY("MSFT",DATE(2025,1,1))','#CONNECT!'),
 ('WEBSERVICE','=WEBSERVICE("https://example.org/")','#CONNECT!'),
]
SKIP={'RTD','STOCKHISTORY','WEBSERVICE'}
def main():
 if os.name!='nt':
  print('This test needs Windows desktop Excel; no parity claims made here.',file=sys.stderr);return 2
 try:import win32com.client as win32
 except ImportError:print('Install pywin32 in this Python environment first.',file=sys.stderr);return 2
 excel=None; wb=None; results=[]
 try:
  excel=win32.DispatchEx('Excel.Application');excel.Visible=False;excel.DisplayAlerts=False
  excel.EnableEvents=False
  wb=excel.Workbooks.Add();ws=wb.Worksheets(1)
  # Avoid modifying Calculation mode: some Excel configurations reject it.
  build=str(excel.Build);version=str(excel.Version)
  for index,(name,formula,gridforge) in enumerate(CASES,1):
   if name in SKIP:
    results.append(dict(name=name,formula=formula,gridforgeOfflineError=gridforge,status='SKIPPED_EXTERNAL',reason='Cannot evaluate without external service or COM server; never automatically connect.'))
    continue
   cell=ws.Cells(index+2,4)
   try:
    cell.Formula2=formula
    excel.CalculateFullRebuild()
    shown=str(cell.Text).strip()
    results.append(dict(name=name,formula=formula,gridforgeOfflineError=gridforge,status=('MATCH_OFFLINE_ERROR' if shown==gridforge else 'DIFFERENT_OFFLINE_ERROR'),excelValue=shown))
   except Exception as exc:
    results.append(dict(name=name,formula=formula,gridforgeOfflineError=gridforge,status='COM_OR_ASSIGNMENT_ERROR',excelError=str(exc)))
  data={'description':'Invalid-connection error comparison only. NOT full parity for any connected function.', 'excelVersion':version,'excelBuild':build,'cases':results,'matchCount':sum(x['status']=='MATCH_OFFLINE_ERROR' for x in results),'skippedExternal':len(SKIP)}
  RESULT.write_text(json.dumps(data,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
  for item in results: print(f"{item['name']}: {item['status']}")
  print('Saved:',RESULT);return 0
 finally:
  if wb is not None:
   try:wb.Close(SaveChanges=False)
   except Exception:pass
  if excel is not None:
   try:excel.Quit()
   except Exception:pass
if __name__=='__main__':sys.exit(main())
