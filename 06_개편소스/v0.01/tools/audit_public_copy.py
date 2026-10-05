"""Compare public RibbonX strings, never VBA/code identity or legal originality."""
from __future__ import annotations
import argparse
import difflib
import hashlib
import json
import re
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

def normalized_label(value: str) -> str:
    text = unicodedata.normalize('NFKC', value.replace('ⓧ', '').replace('Ⓧ', ''))
    text = re.sub(r'\s*[\[(]?(?:Ctrl|Alt|Shift)\s*[+＋].*$', '', text, flags=re.I)
    text = re.sub(r'\s+F(?:1[0-2]|[1-9])\s*$', '', text, flags=re.I)
    return re.sub(r'\s+', '', text).casefold()

def ribbon_rows(path: Path) -> list[dict]:
    if path.suffix.lower() == '.xml':
        payload = path.read_bytes()
    else:
        with zipfile.ZipFile(path) as package:
            names = [name for name in package.namelist() if name.lower() in {'customui/customui14.xml','customui/customui.xml'}]
            if not names:
                raise ValueError('RibbonX not found')
            selected = next((name for name in names if name.lower().endswith('customui14.xml')),names[0])
            payload = package.read(selected)
    rows = []
    for element in ET.fromstring(payload).iter():
        if not element.get('onAction') or not element.get('label'):
            continue
        rows.append({'id':element.get('id') or element.get('idMso') or '',
                     'label':element.get('label',''),
                     'description':element.get('supertip') or element.get('screentip','')})
    return rows

def compare(rows: list[dict], references: list[dict]) -> dict:
    exact = []
    normalized = []
    help_overlap = []
    for row in rows:
        names = [ref['label'] for ref in references if row['label'] == ref['label']]
        if names:
            exact.append({'id':row['id'],'label':row['label'],'reference_labels':sorted(set(names))})
        names = [ref['label'] for ref in references if normalized_label(row['label']) == normalized_label(ref['label'])]
        if names:
            normalized.append({'id':row['id'],'label':row['label'],'reference_labels':sorted(set(names))})
        text = re.sub(r'\s+', '', row['description'])
        if len(text) < 30:
            continue
        best = None
        for ref in references:
            other = re.sub(r'\s+', '',ref['description'])
            if len(other) < 30:
                continue
            matcher = difflib.SequenceMatcher(None,text,other,autojunk=False)
            block = matcher.find_longest_match()
            ratio = matcher.ratio()
            if ratio >= .72 or block.size >= 30:
                candidate = {'id':row['id'],'label':row['label'],'reference_label':ref['label'],
                             'similarity':round(ratio,4),'longest_shared_characters':block.size,
                             'shared_fragment':text[block.a:block.a+block.size]}
                if best is None or (ratio,block.size) > (best['similarity'],best['longest_shared_characters']):
                    best = candidate
        if best:
            help_overlap.append(best)
    return {'execution_controls':len(rows),'exact_label_count':len(exact),
            'normalized_label_count':len(normalized),'help_review_candidate_count':len(help_overlap),
            'exact_labels':exact,'normalized_labels':normalized,'help_review_candidates':help_overlap}

def build_report(before: Path, after: Path, reference: Path) -> dict:
    refs = ribbon_rows(reference)
    return {'evidence_class':'STATIC_RIBBONX_TEXT_ONLY','method':{
        'labels':'NFKC, whitespace/case, trailing Ctrl/Alt/Shift/F1-F12 hints; not semantic identity',
        'help':'whitespace removed; similarity >= 0.72 or contiguous overlap >= 30 characters; manual review candidates only',
        'limits':'Counts are controls, including repeated surface placements. Generic Excel terms can match. No code, behavior or legal originality verdict.'},
        'inputs':{key:{'path':str(path),'sha256':hashlib.sha256(path.read_bytes()).hexdigest()} for key,path in [('before',before),('after',after),('reference',reference)]},
        'reference_execution_controls':len(refs),'before':compare(ribbon_rows(before),refs),'after':compare(ribbon_rows(after),refs)}

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    for name in ('before','after','reference','output'):
        parser.add_argument('--'+name,type=Path,required=True)
    args=parser.parse_args()
    if args.output.exists():
        raise SystemExit('Refusing to overwrite existing audit evidence')
    report=build_report(args.before,args.after,args.reference)
    args.output.parent.mkdir(parents=True,exist_ok=True)
    args.output.write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps({key:{field:report[key][field] for field in ('execution_controls','exact_label_count','normalized_label_count','help_review_candidate_count')} for key in ('before','after')},ensure_ascii=False))

if __name__ == '__main__':
    main()
