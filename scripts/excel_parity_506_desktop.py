#!/usr/bin/env python3
"""Evaluate one fixture per registered GridForge function in *actual* Windows Excel.

Requires Python + pywin32 + Microsoft Excel desktop. Does not open user files,
execute VBA, save workbook, access external files, or change Calculation mode.
A local, blank scratch workbook is used. Unsupported functions are recorded
individually rather than aborting the 506-case audit.
"""
import argparse
import hashlib
import json
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FIXTURE = ROOT / "tests" / "excel_parity_506_cases.json"
RESULT = ROOT / "docs" / "PARITY_506_EXCEL_RESULTS.json"
ONLINE_FUNCTIONS = {'IMAGE', 'TRANSLATE', 'DETECTLANGUAGE'}
ERROR_CODES = ('#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#NULL!',
               '#SPILL!', '#CALC!', '#FIELD!', '#BLOCKED!', '#GETTING_DATA', '#CONNECT!',
               '#BUSY!', '#UNKNOWN!', '#PYTHON!', '#CIRC!')


def specimen_hash():
    return hashlib.sha256(FIXTURE.read_bytes()).hexdigest()


def normalize(value, cell):
    try:
        shown = str(cell.Text or '').strip()
    except Exception:
        shown = ''
    for code in ERROR_CODES:
        if shown == code or (shown.startswith(code) and code == '#N/A'):
            return {'type': 'error', 'value': code}
    if value is None:
        return {'type': 'blank', 'value': None}
    if isinstance(value, bool):
        return {'type': 'boolean', 'value': value}
    if isinstance(value, (int, float)):
        return {'type': 'number', 'value': float(value)}
    if isinstance(value, str):
        return {'type': 'string', 'value': value}
    return {'type': 'other', 'value': str(value)}


def save(cases, session, complete):
    result = {**session, 'complete': complete, 'done': len(cases), 'cases': cases}
    RESULT.parent.mkdir(parents=True, exist_ok=True)
    temp = RESULT.with_suffix('.json.tmp')
    temp.write_text(json.dumps(result, indent=2, ensure_ascii=False) + '\n', encoding='utf-8')
    temp.replace(RESULT)


def main():
    parser = argparse.ArgumentParser(description='Actual Excel audit of all GridForge registered functions')
    parser.add_argument('--limit', type=int, default=0, help='Run first N cases as a smoke test (0=all)')
    parser.add_argument('--progress-every', type=int, default=25)
    parser.add_argument('--allow-online-functions', action='store_true',
                        help='Opt in to cloud-connected functions (may send text to Microsoft)')
    args = parser.parse_args()
    if os.name != 'nt':
        print('Windows desktop Excel is required. GridForge-only fixture results are already bundled.', file=sys.stderr)
        return 2
    try:
        import win32com.client as win32
    except ImportError:
        print('pywin32 missing in current interpreter. Run: python -m pip install pywin32', file=sys.stderr)
        return 2
    raw = json.loads(FIXTURE.read_text(encoding='utf-8'))
    cases = raw['cases']
    if len(cases) != 506 or len({c['function'] for c in cases}) != 506:
        print('STOP: Fixture registry coverage is not 506 unique functions.', file=sys.stderr)
        return 2
    target = cases[:args.limit] if args.limit > 0 else cases
    if args.limit < 0 or not target:
        print('Invalid case limit.', file=sys.stderr)
        return 2
    # Avoid mixing a partial smoke test with a previous full run.
    actual_start = time.monotonic()
    app = book = None
    outcome = []
    session = {'schema': 'gridforge-excel-full-result/1', 'source': 'Microsoft Excel desktop',
               'excelExecuted': True, 'specimenSHA256': specimen_hash(),
               'fixtureCount': len(cases), 'requestedCount': len(target)}
    try:
        app = win32.DispatchEx('Excel.Application')
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.AskToUpdateLinks = False
        app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        # Intentionally DO NOT set app.Calculation: setting it before a workbook
        # existed caused COM failure on Excel 16.0 build 20430.
        book = app.Workbooks.Add()
        sheet = book.Worksheets(1)
        sheet.Columns('Z').ColumnWidth = 55
        session.update(excelVersion=str(app.Version), excelBuild=str(app.Build), complete=False)
        print(f'Connected to real Excel {session["excelVersion"]}, build {session["excelBuild"]}.')
        for i, case in enumerate(target, 1):
            formula = case['formula']
            row = {'id': case['id'], 'function': case['function'], 'formula': formula}
            if case['function'] in ONLINE_FUNCTIONS and not args.allow_online_functions:
                row['excel'] = {'type': 'online-function-skipped',
                                'value': 'Opt in using --allow-online-functions; may contact online services.'}
                outcome.append(row)
                if i % max(1,args.progress_every) == 0 or i == len(target):
                    session['elapsedSeconds'] = round(time.monotonic()-actual_start, 2)
                    save(outcome, session, i==len(cases))
                    print(f'Excel reviewed {i}/{len(target)} | elapsed {session["elapsedSeconds"]}s', flush=True)
                continue
            try:
                # Each case begins with an empty spreadsheet: no cross-case leakage.
                sheet.Range('A1:AZ45').ClearContents()
                sheet.Range('A1:AZ45').NumberFormat = 'General'
                for address, raw_value in case.get('inputs', {}).items():
                    cell = sheet.Range(address)
                    if isinstance(raw_value, dict) and 'formula' in raw_value:
                        cell.Formula2 = raw_value['formula']
                    elif isinstance(raw_value, str):
                        cell.NumberFormat = '@'
                        cell.Value2 = raw_value
                    else:
                        cell.Value2 = raw_value
                cell = sheet.Range(raw['comparisonCell'])
                try:
                    cell.Formula2 = formula
                    row['excelInputMode'] = 'Formula2'
                except Exception as exc:
                    try:
                        cell.Formula = formula
                        row['excelInputMode'] = 'Formula'
                    except Exception as fallback_exc:
                        row['excel'] = {'type': 'formula-assignment-error',
                                        'value': str(fallback_exc), 'formula2Error': str(exc)}
                if 'excel' not in row:
                    try:
                        app.Calculate()
                        row['excel'] = normalize(cell.Value2, cell)
                    except Exception as exc:
                        row['excel'] = {'type': 'excel-calculation-error', 'value': str(exc)}
            except Exception as exc:
                row['excel'] = {'type': 'excel-case-error', 'value': str(exc)}
            outcome.append(row)
            if i % max(1,args.progress_every) == 0 or i == len(target):
                session['elapsedSeconds'] = round(time.monotonic()-actual_start, 2)
                save(outcome, session, i==len(cases))
                errors=sum(x['excel']['type'].endswith('-error') for x in outcome)
                print(f'Excel evaluated {i}/{len(target)} | COM failures {errors} | elapsed {session["elapsedSeconds"]}s', flush=True)
        print(f'Finished: {len(outcome)} Excel formulas. Saved {RESULT}.')
        if args.limit > 0 and len(target)<len(cases):
            print('SMOKE TEST ONLY. Rerun without --limit for the full 506-function report.')
        return 0
    except Exception as exc:
        print('Excel execution interrupted:', repr(exc), file=sys.stderr)
        if outcome:
            session['elapsedSeconds'] = round(time.monotonic()-actual_start, 2)
            save(outcome,session,False)
        return 1
    finally:
        if book is not None:
            try: book.Close(SaveChanges=False)
            except Exception: pass
        if app is not None:
            try: app.Quit()
            except Exception: pass


if __name__ == '__main__':
    raise SystemExit(main())
