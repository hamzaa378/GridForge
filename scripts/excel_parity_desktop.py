#!/usr/bin/env python3
"""Windows-only, actual Microsoft Excel calculation reference for GridForge fixtures.

No user workbook is opened. No VBA is executed. All workbooks remain unsaved.
Requires desktop Excel and `py -3 -m pip install pywin32`.
"""
import json, os, sys, time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
INPUT = ROOT / 'tests' / 'excel_parity_cases.json'
OUTPUT = ROOT / 'docs' / 'PARITY_EXCEL_RESULTS.json'


def normalize(value, cell):
    # Excel COM returns a COM error Variant for spreadsheet errors; .Text reliably exposes the code.
    shown = str(cell.Text or '')
    if shown.startswith(('#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#NULL!', '#SPILL!', '#CALC!', '#FIELD!', '#BLOCKED!')):
        return {'type': 'error', 'value': shown}
    if value is None:
        return {'type': 'blank', 'value': None}
    if isinstance(value, bool):
        return {'type': 'boolean', 'value': value}
    if isinstance(value, (float, int)):
        return {'type': 'number', 'value': float(value)}
    if isinstance(value, str):
        return {'type': 'string', 'value': value}
    return {'type': 'other', 'value': str(value)}


def main():
    if os.name != 'nt':
        print('NOT EXECUTED: actual Microsoft Excel COM automation is Windows-only.', file=sys.stderr)
        return 2
    try:
        import win32com.client as win32
    except ImportError:
        print('Please install pywin32: py -3 -m pip install pywin32', file=sys.stderr)
        return 2
    suite = json.loads(INPUT.read_text(encoding='utf-8'))
    excel = None
    wb = None
    started = time.monotonic()
    try:
        excel = win32.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        wb = excel.Workbooks.Add()
        first = wb.Worksheets(1)
        sheets = []
        for i, case in enumerate(suite['cases']):
            sheet = first if i == 0 else wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            sheet.Name = 'P%04d' % (i + 1)
            for ref, value in case.get('inputs', {}).items():
                cell = sheet.Range(ref)
                if isinstance(value, dict) and 'text' in value:
                    cell.NumberFormat = '@'
                    cell.Value2 = value['text']
                else:
                    cell.Value2 = value
            result_cell = sheet.Range(suite['comparisonCell'])
            try:
                result_cell.Formula2 = case['formula']
            except Exception:
                result_cell.Formula = case['formula']
            sheets.append((case, sheet))
        excel.CalculateFullRebuild()
        outcomes = []
        for case, sheet in sheets:
            cell = sheet.Range(suite['comparisonCell'])
            try:
                result = normalize(cell.Value2, cell)
            except Exception as exc:
                result = {'type': 'excel-com-error', 'value': str(exc)}
            outcomes.append({'id': case['id'], 'excel': result})
        report = {
            'schema': 'gridforge-excel-parity-result/1', 'engine': 'Microsoft Excel desktop',
            'excelExecuted': True, 'excelVersion': str(excel.Version),
            'excelBuild': str(excel.Build) if hasattr(excel, 'Build') else 'unknown',
            'elapsedSeconds': round(time.monotonic()-started, 2), 'cases': outcomes
        }
        OUTPUT.write_text(json.dumps(report, indent=2, ensure_ascii=False) + '\n', encoding='utf-8')
        print('ACTUAL EXCEL EXECUTION COMPLETE: %d cases, Excel %s, %.2fs' % (len(outcomes), report['excelVersion'], report['elapsedSeconds']))
        print('Saved:', OUTPUT)
        return 0
    except Exception as exc:
        print('EXCEL EXECUTION FAILED:', type(exc).__name__, exc, file=sys.stderr)
        return 1
    finally:
        if wb is not None:
            try: wb.Close(SaveChanges=False)
            except Exception: pass
        if excel is not None:
            try: excel.Quit()
            except Exception: pass

if __name__ == '__main__':
    sys.exit(main())
