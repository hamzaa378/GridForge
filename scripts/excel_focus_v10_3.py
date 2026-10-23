#!/usr/bin/env python3
"""Focused recheck for v10.3 residual parity findings, using desktop Excel.
Requires Windows, Microsoft Excel, pywin32. Never opens existing workbooks,
changes Excel security settings, writes user content, or enables macros.
"""
import json
import os
import sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent.parent
TARGET=ROOT/'docs'/'EXCEL_FOCUSED_RECHECK.json'
CASES=[
 ('IMPOWER','Z20','=IMPOWER("1+i","2")',{}),
 ('XNPV_fixture','Z20','=XNPV(0,F1:F5,G1:G5)',{
  **{f'F{i}':v for i,v in enumerate([-1000,300,400,500,200],1)},
  **{f'G{i}':{'formula':f'=DATE({year},1,1)'} for i,year in enumerate([2020,2021,2022,2023,2024],1)}}),
 ('XNPV_serials','Z20','=XNPV(0,F1:F5,G1:G5)',{
  **{f'F{i}':v for i,v in enumerate([-1000,300,400,500,200],1)},
  **{f'G{i}':v for i,v in enumerate([43831,44197,44562,44927,45292],1)}}),
]
def main():
 if os.name!='nt':
  print('Desktop Excel on Windows required. No result generated.');return 2
 try:import win32com.client as win32
 except ImportError:
  print('Install pywin32 in your active Python environment.');return 2
 app=None;wb=None;out=[]
 try:
  app=win32.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False;app.AutomationSecurity=3
  wb=app.Workbooks.Add();ws=wb.Worksheets(1)
  for name,target,formula,inputs in CASES:
   ws.Range('A1:AZ45').ClearContents();ws.Range('A1:AZ45').NumberFormat='General'
   for ref,value in inputs.items():
    if isinstance(value,dict):ws.Range(ref).Formula2=value['formula']
    else:ws.Range(ref).Value2=value
   row={'case':name,'formula':formula,'excelVersion':str(app.Version)}
   try:
    ws.Range(target).Formula2=formula;app.CalculateFullRebuild()
    v=ws.Range(target).Value2;shown=str(ws.Range(target).Text)
    row.update(value=v,display=shown,inputs={k:(v['formula'] if isinstance(v,dict) else v) for k,v in inputs.items()})
   except Exception as e:row.update(error=str(e))
   print(name, row.get('display',row.get('error')));out.append(row)
  TARGET.write_text(json.dumps({'scope':'Focused Excel live recheck; not full 506-run','cases':out},indent=2)+'\n',encoding='utf8')
  print('Saved',TARGET);return 0
 finally:
  if wb is not None:
   try:wb.Close(SaveChanges=False)
   except Exception:pass
  if app is not None:
   try:app.Quit()
   except Exception:pass
if __name__=='__main__':sys.exit(main())
