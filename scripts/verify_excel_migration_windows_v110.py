#!/usr/bin/env python3
"""Real desktop Excel acceptance: open the GridForge v11.0 OOXML demo safely, read-only.
Windows-only, pywin32 and a licensed Microsoft Excel installation required.
No VBA, external links, or online functions are run by the verifier.
"""
import json
from pathlib import Path
ROOT=Path(__file__).resolve().parent.parent
SAMPLE=ROOT/'examples'/'GridForge_v11_0_Table_Migration_Demo.xlsx'
OUT=ROOT/'docs'/'EXCEL_MIGRATION_WINDOWS_V11_0_RESULTS.json'

def main():
    import win32com.client
    excel=win32com.client.DispatchEx('Excel.Application')
    excel.Visible=False
    excel.DisplayAlerts=False
    book=None
    try:
        excel.AutomationSecurity=3
        book=excel.Workbooks.Open(str(SAMPLE.resolve()),UpdateLinks=0,ReadOnly=True)
        sheet=book.Worksheets(1)
        values={cell:sheet.Range(cell).Value for cell in ('F2','D7','B7')}
        tables=int(sheet.ListObjects.Count)
        charts=int(sheet.ChartObjects().Count)
        # Both values are generated from local, deterministic workbook data.
        expected=3085
        passed=tables>=1 and charts>=2 and isinstance(values['F2'],(int,float)) and abs(values['F2']-expected)<1e-7
        report={
          'scope':'Read-only Excel desktop verification of GridForge v11.0 native table/chart and structured-reference formula; not complete Excel compatibility',
          'excelVersion':excel.Version,'excelBuild':excel.Build,
          'tableObjects':tables,'chartObjects':charts,'cells':values,
          'totalExpected':expected,'passed':bool(passed)
        }
        OUT.write_text(json.dumps(report,indent=2,default=str)+'\n',encoding='utf-8')
        print(json.dumps(report,indent=2,default=str))
        if not passed:raise SystemExit('Excel acceptance failed: investigate OOXML or calculation differences.')
    finally:
        if book is not None:book.Close(SaveChanges=False)
        excel.Quit()
if __name__=='__main__':main()
