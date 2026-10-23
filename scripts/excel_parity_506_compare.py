#!/usr/bin/env python3
"""Truthful Excel-vs-GridForge comparison: reports every function, including gaps."""
import hashlib
import json
import math
import sys
from collections import Counter
from pathlib import Path

ROOT=Path(__file__).resolve().parent.parent
DOC=ROOT/'docs'
FIXTURE=ROOT/'tests/excel_parity_506_cases.json'
XLS=DOC/'PARITY_506_EXCEL_RESULTS.json'
GF=DOC/'PARITY_506_GRIDFORGE_RESULTS.json'
OUTPUT=DOC/'PARITY_506_COMPARISON.json'
MD=DOC/'PARITY_506_COMPARISON.md'


def equivalent(a,b):
    if not a or not b or a['type'] != b['type']:
        return False
    if a['type']=='number':
        # Excel's analysis functions use approximate solvers; preserve tolerance
        # explicitly rather than rounding both values into forced equality.
        return math.isclose(a['value'],b['value'],rel_tol=1e-7,abs_tol=1e-8)
    return a.get('value') == b.get('value')


def compare():
    if not XLS.exists():
        raise RuntimeError('Actual Excel evidence missing; run python scripts/excel_parity_506_desktop.py on Windows')
    fixture=json.loads(FIXTURE.read_text('utf-8'))
    g=json.loads(GF.read_text('utf-8'))
    x=json.loads(XLS.read_text('utf-8'))
    expected=hashlib.sha256(FIXTURE.read_bytes()).hexdigest()
    if not x.get('excelExecuted') or x.get('source')!='Microsoft Excel desktop' or x.get('specimenSHA256')!=expected:
        raise RuntimeError('Excel evidence missing, unverified, or from a different fixture version')
    if not x.get('complete') or x.get('done') != 506 or len(x.get('cases',[]))!=506:
        raise RuntimeError(f'Incomplete Excel run ({x.get("done",0)}/506). Run full suite without --limit.')
    if len(fixture['cases'])!=506 or len(g.get('cases',[]))!=506:
        raise RuntimeError('Fixture or GridForge result does not cover 506 functions')
    by_g={r['function']:r for r in g['cases']}
    by_x={r['function']:r for r in x['cases']}
    if len(by_g)!=506 or len(by_x)!=506:
        raise RuntimeError('Duplicate or missing function results')
    rows=[]
    for c in fixture['cases']:
        name=c['function']
        a=by_g[name]['gridforge']
        b=by_x[name]['excel']
        if by_g[name]['formula']!=c['formula'] or by_x[name]['formula']!=c['formula']:
            raise RuntimeError(f'Stale or mismatched formula for {name}')
        scope=c['scope']
        if b['type']=='online-function-skipped':
            status='ONLINE_FUNCTION_NOT_RUN'
        elif b['type'].endswith('-error'):
            status='EXCEL_EXECUTION_FAILURE'
        elif b['type']=='error' and b['value'] in ('#NAME?','#BLOCKED!') and a['type'] not in ('error','exception'):
            status='EXCEL_UNAVAILABLE_OR_REJECTED'
        elif scope=='volatile':
            status='VOLATILE_NONDETERMINISTIC'
        elif equivalent(a,b):
            status='MATCH'
        elif scope=='limited-semantics':
            status='KNOWN_SEMANTIC_DIFFERENCE'
        else:
            status='DIFFERENCE'
        rows.append({'function':name,'scope':scope,'formula':c['formula'],'status':status,'gridforge':a,'excel':b})
    counts=dict(Counter(r['status'] for r in rows))
    ordered=['MATCH','DIFFERENCE','KNOWN_SEMANTIC_DIFFERENCE','EXCEL_UNAVAILABLE_OR_REJECTED','EXCEL_EXECUTION_FAILURE','ONLINE_FUNCTION_NOT_RUN','VOLATILE_NONDETERMINISTIC']
    counts={name:counts.get(name,0) for name in ordered}
    comparable=counts['MATCH']+counts['DIFFERENCE']+counts['KNOWN_SEMANTIC_DIFFERENCE']
    summary={'excelExecuted':True,'excelVersion':x['excelVersion'],'excelBuild':x['excelBuild'],
             'registeredFunctions':506,'functionFixturesExecuted':506,
             'oneCasePerFunction':True,'counts':counts,
             'comparableFunctions':comparable,
             'agreementPercent':round(100*counts['MATCH']/comparable,2) if comparable else None,
             'numericTolerance':'absolute 1e-8, relative 1e-7',
             'caveat':'One fixture per name; agreement does not imply complete Excel semantic or workbook compatibility.'}
    OUTPUT.write_text(json.dumps({'summary':summary,'functions':rows},indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
    lines=['# GridForge 506-function comparison with desktop Excel','',
       f"Excel {x['excelVersion']} (build {x['excelBuild']}); **{comparable} comparable cases**, **{counts['MATCH']} matches**, **{counts['DIFFERENCE']+counts['KNOWN_SEMANTIC_DIFFERENCE']} differences**.",'',
       f"Comparable-case agreement: **{summary['agreementPercent']}%**. This is one invocation per registered name, not proof of general Excel compatibility.",'', '## Outcome summary','']
    lines.extend(f'- **{key}**: {counts[key]}' for key in ordered)
    lines += ['','## Differences and limitations','']
    for r in rows:
        if r['status']!='MATCH':
            lines.append(f"- **{r['function']}** [{r['status']}]: `{r['formula']}`  GridForge: `{r['gridforge']}`; Excel: `{r['excel']}`")
    lines+=['','## Methodology and caveats','','Each case invokes the named function in an isolated Excel worksheet state. Some matrix results are read with INDEX rather than testing automatic spill. Time-dependent functions are excluded from exact equality; Excel-version availability is reported separately. Other inputs, locales, error cases, workbooks, external links, macros, and dynamic-array behavior remain unverified by these 506 fixtures.']
    MD.write_text('\n'.join(lines)+'\n',encoding='utf-8')
    return summary

if __name__=='__main__':
    try:
        print(json.dumps(compare(),indent=2))
        print('Saved',OUTPUT,'and',MD)
    except Exception as exc:
        print('NOT A VALID 506-FUNCTION EXCEL COMPARISON:',exc,file=sys.stderr)
        raise SystemExit(2)
