#!/usr/bin/env python3
"""Real Microsoft Excel smoke test for v10.8 OOXML chart interoperability.
Does not run any VBA; opens a bundled synthetic workbook read-only and closes without saving.
Requires Windows, desktop Excel, and pywin32 in the active Python environment.
"""
from __future__ import annotations
import json, pathlib, sys, time
ROOT = pathlib.Path(__file__).resolve().parent.parent
SOURCE = ROOT / 'examples' / 'native_charts_v108.xlsx'
OUT = ROOT / 'docs' / 'V10_8_REAL_EXCEL_CHART_RESULTS.json'

def main():
    if sys.platform != 'win32':
        print('Needs Windows with Microsoft Excel installed. No Excel claim was made.');return 2
    if not SOURCE.is_file():
        print('Missing chart fixture: ' + str(SOURCE));return 2
    try:
        import win32com.client
    except ImportError:
        print('Activate your Windows environment, then: python -m pip install pywin32');return 2
    excel=None;book=None;result={'verified':False,'fixture':str(SOURCE),'timestamp':time.strftime('%Y-%m-%dT%H:%M:%S'),'expectedChartCount':9}
    try:
        excel=win32com.client.DispatchEx('Excel.Application')
        excel.Visible=False;excel.DisplayAlerts=False
        result['excelVersion']=str(excel.Version)
        book=excel.Workbooks.Open(str(SOURCE),UpdateLinks=0,ReadOnly=True)
        charts=[]
        for sheet in book.Worksheets:
            objects=sheet.ChartObjects()
            for index in range(1,int(objects.Count)+1):
                obj=objects.Item(index)
                charts.append({'sheet':str(sheet.Name),'objectName':str(obj.Name),'seriesCount':int(obj.Chart.SeriesCollection().Count),'chartType':int(obj.Chart.ChartType)})
        result['charts']=charts;result['observedChartCount']=len(charts)
        result['verified']=len(charts)==9 and all(ch['seriesCount']>=1 for ch in charts)
    except Exception as exc:
        result['error']=str(exc)
    finally:
        if book is not None:
            try:book.Close(SaveChanges=False)
            except Exception:pass
        if excel is not None:
            try:excel.Quit()
            except Exception:pass
        OUT.write_text(json.dumps(result,indent=2),encoding='utf-8')
    print(json.dumps(result,indent=2))
    print('Report:',OUT)
    return 0 if result['verified'] else 1
if __name__=='__main__':raise SystemExit(main())
