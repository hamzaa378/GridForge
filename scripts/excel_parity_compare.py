#!/usr/bin/env python3
"""Python-only comparison of actual Excel output and precomputed GridForge results."""
import json, math, sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent.parent
DOC=ROOT/'docs'

def load(name): return json.loads((DOC/name).read_text(encoding='utf-8'))

def same(a,b):
    if not a or not b: return False
    if a['type']=='number' and b['type']=='number':
        return math.isclose(a['value'],b['value'],rel_tol=1e-8,abs_tol=1e-8)
    return a['type']==b['type'] and a['value']==b['value']

def main():
    target=DOC/'PARITY_EXCEL_RESULTS.json'
    if not target.exists():
        print('NOT MEASURED. Run: py -3 scripts\\excel_parity_desktop.py (Windows + desktop Excel + pywin32).')
        return 2
    xl=load('PARITY_EXCEL_RESULTS.json')
    if xl.get('excelExecuted') is not True or xl.get('engine')!='Microsoft Excel desktop':
        print('Refusing to present an unverified report as Excel results.')
        return 2
    engine=load('PARITY_GRIDFORGE_RESULTS.json')
    cases=json.loads((ROOT/'tests'/'excel_parity_cases.json').read_text(encoding='utf-8'))['cases']
    excel={c['id']:c['excel'] for c in xl['cases']}
    gridforge={c['id']:c['gridforge'] for c in engine['cases']}
    rows=[]
    for c in cases:
        a=gridforge.get(c['id']); b=excel.get(c['id']); kind=c.get('kind','scalar')
        if not a or not b: status='MISSING_EVIDENCE'
        elif kind=='expected-gap': status='KNOWN_LIMITATION'
        elif b['type']=='error' and b['value'] in ('#NAME?','#BLOCKED!'): status='EXCEL_VERSION_OR_SECURITY_UNSUPPORTED'
        elif same(a,b): status='MATCH'
        else: status='DIFFERENCE'
        rows.append({'id':c['id'],'formula':c['formula'],'category':c['category'],'status':status,'gridforge':a,'excel':b})
    labels=['MATCH','DIFFERENCE','KNOWN_LIMITATION','EXCEL_VERSION_OR_SECURITY_UNSUPPORTED','MISSING_EVIDENCE']
    counts={label:sum(x['status']==label for x in rows) for label in labels}
    n=counts['MATCH']+counts['DIFFERENCE']
    summary={'actualExcelExecuted':True,'excelVersion':xl['excelVersion'],'excelBuild':xl.get('excelBuild'),'tests':len(rows),'counts':counts,'scalarCompared':n,'scalarAgreementPercent':round(100*counts['MATCH']/n,2) if n else None}
    (DOC/'PARITY_COMPARISON.json').write_text(json.dumps({'summary':summary,'cases':rows},indent=2)+'\n',encoding='utf-8')
    lines=['# GridForge vs Microsoft Excel: actual Windows comparison','',f"Excel {xl['excelVersion']}; {n} comparable scalar cases; **{counts['MATCH']} match, {counts['DIFFERENCE']} differ**.",'',f"Comparable-case agreement: **{summary['scalarAgreementPercent']}%** (not a percentage of all Excel capabilities).",'', '## Outcome categories','']
    lines += [f"- {label}: {counts[label]}" for label in labels]
    lines += ['','## Mismatches','']
    lines += [f"- {r['id']} `{r['formula']}`: GridForge {r['gridforge']}, Excel {r['excel']}" for r in rows if r['status']=='DIFFERENCE']
    lines += ['','## Interpretation','','Actual test results apply only to this bounded deterministic fixture set, Excel version, and operating environment. This is NOT proof of complete workbook, VBA, pivot, chart, spill, locale, or file-format parity.']
    (DOC/'PARITY_COMPARISON.md').write_text('\n'.join(lines)+'\n',encoding='utf-8')
    print(json.dumps(summary,indent=2))
    print('Wrote docs/PARITY_COMPARISON.md and docs/PARITY_COMPARISON.json')
    return 1 if counts['DIFFERENCE'] else 0

if __name__=='__main__':sys.exit(main())
