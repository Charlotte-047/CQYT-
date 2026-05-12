from __future__ import annotations
from zipfile import ZipFile
from pathlib import Path
from lxml import etree as ET
import re, sys

W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS={'w':W.strip('{}')}
def q(t): return f'{W}{t}'
def text(p): return ''.join(t.text or '' for t in p.findall('.//w:t',NS)).strip()
def jc(p):
    ppr=p.find(q('pPr')); n=None if ppr is None else ppr.find(q('jc'))
    return '' if n is None else (n.get(q('val')) or '')
def spacing(p):
    ppr=p.find(q('pPr')); n=None if ppr is None else ppr.find(q('spacing'))
    return {} if n is None else {k.split('}')[-1]:v for k,v in n.attrib.items()}
def ind(p):
    ppr=p.find(q('pPr')); n=None if ppr is None else ppr.find(q('ind'))
    return {} if n is None else {k.split('}')[-1]:v for k,v in n.attrib.items()}
def flags(p):
    ppr=p.find(q('pPr'))
    return set() if ppr is None else {c.tag.split('}')[-1] for c in ppr}
def run_props(p):
    vals=[]
    for r in p.findall('./w:r',NS):
        rpr=r.find(q('rPr')); rf=None if rpr is None else rpr.find(q('rFonts')); sz=None if rpr is None else rpr.find(q('sz'))
        b=None if rpr is None else rpr.find(q('b'))
        vals.append({'east':None if rf is None else rf.get(q('eastAsia')),'ascii':None if rf is None else rf.get(q('ascii')),'hAnsi':None if rf is None else rf.get(q('hAnsi')),'size':None if sz is None else sz.get(q('val')),'bold':(b is not None and b.get(q('val')) not in ('0','false','False'))})
    return vals
def border_map(el):
    out={}
    b=el.find(q('tblBorders')) if el is not None else None
    if b is None: return out
    for name in ('top','left','bottom','right','insideH','insideV'):
        n=b.find(q(name))
        if n is not None: out[name]={k.split('}')[-1]:v for k,v in n.attrib.items()}
    return out
def cell_border_map(tc):
    tcPr=tc.find(q('tcPr')); b=None if tcPr is None else tcPr.find(q('tcBorders'))
    out={}
    if b is None: return out
    for name in ('top','left','bottom','right'):
        n=b.find(q(name))
        if n is not None: out[name]={k.split('}')[-1]:v for k,v in n.attrib.items()}
    return out
def is_table_caption(t): return bool(re.match(r'^表\s*\d+(?:\.\d+)*\s+\S+', t or ''))
def is_blank(p): return text(p)=='' and len(p.findall('.//w:drawing',NS))==0
def border_is_none(b): return (not b) or b.get('val') in (None,'nil','none') or b.get('sz')=='0'
def main(doc):
    p=Path(doc); fails=[]; warns=[]
    with ZipFile(p) as z:
        root=ET.fromstring(z.read('word/document.xml'))
    body=root.find(q('body')); children=list(body); tables=body.findall('.//w:tbl',NS)
    tbl_positions=[i for i,n in enumerate(children) if n.tag==q('tbl')]
    print('TABLE_COUNT',len(tbl_positions))
    for ti,idx in enumerate(tbl_positions,1):
        tbl=children[idx]
        rows=tbl.findall('./w:tr',NS)
        # Caption must be immediately above table, no blank between.
        j=idx-1
        blanks=0
        while j>=0 and children[j].tag==q('p') and is_blank(children[j]): blanks+=1; j-=1
        if blanks: fails.append(f'table {ti} has blank between caption and table: blanks={blanks}')
        if j<0 or children[j].tag!=q('p') or not is_table_caption(text(children[j])):
            prev=[text(children[k]) for k in range(max(0,idx-4),idx) if children[k].tag==q('p')]
            fails.append(f'table {ti} not immediately preceded by table caption; prev={prev}')
        else:
            cap=children[j]; t=text(cap)
            if jc(cap)!='center': fails.append(f'table {ti} caption not centered: {t} jc={jc(cap)}')
            sp=spacing(cap)
            if sp.get('line')!='440' or sp.get('lineRule')!='exact': fails.append(f'table {ti} caption line spacing not fixed 22pt: {sp}')
            if 'keepNext' not in flags(cap) or 'keepLines' not in flags(cap): fails.append(f'table {ti} caption missing keepNext/keepLines: {flags(cap)}')
            if re.search(r'[。；;,.，]$',t): fails.append(f'table {ti} caption ends with punctuation: {t}')
            for rp in [x for x in run_props(cap) if any(x.values())]:
                if rp.get('east')!='宋体' or rp.get('ascii')!='Times New Roman' or rp.get('hAnsi')!='Times New Roman' or rp.get('size')!='21':
                    fails.append(f'table {ti} caption font wrong: {t} props={rp}')
        # Table no page split: every row cantSplit; paragraphs keepLines and row-bridging keepNext except last row.
        for ri,row in enumerate(rows):
            trPr=row.find(q('trPr'))
            if trPr is None or trPr.find(q('cantSplit')) is None: fails.append(f'table {ti} row {ri+1} missing cantSplit')
            for pa in row.findall('.//w:p',NS):
                if text(pa):
                    fl=flags(pa)
                    if 'keepLines' not in fl: fails.append(f'table {ti} row {ri+1} paragraph missing keepLines: {text(pa)[:30]}')
                    if ri < len(rows)-1 and 'keepNext' not in fl: fails.append(f'table {ti} row {ri+1} paragraph missing keepNext for no-split: {text(pa)[:30]}')
        # Strict three-line per standard t0: visible lines are cell borders.
        # - header row cell top = 1.5pt (w:sz 12)
        # - header row cell bottom = 0.75pt (w:sz 6)
        # - last row cell bottom = 1.5pt (w:sz 12)
        # - no left/right/internal/body row borders
        tblPr=tbl.find(q('tblPr')); bm=border_map(tblPr.find(q('tblBorders')) if tblPr is not None else None)
        for side,b in bm.items():
            if not border_is_none(b): fails.append(f'table {ti} unexpected tbl {side} border; standard t0 uses cell borders: {b}')
        if rows:
            for ri,row in enumerate(rows):
                for ci,tc in enumerate(row.findall('./w:tc',NS),1):
                    cb=cell_border_map(tc)
                    for side in ('left','right','insideH','insideV','tl2br','tr2bl'):
                        if not border_is_none(cb.get(side,{})): fails.append(f'table {ti} cell {ri+1},{ci} unexpected {side} border: {cb.get(side)}')
                    top=cb.get('top',{})
                    bottom=cb.get('bottom',{})
                    if ri==0:
                        if top.get('val')!='single' or top.get('sz')!='12': fails.append(f'table {ti} header cell {ci} top not 1.5pt: {top}')
                        if bottom.get('val')!='single' or bottom.get('sz')!='6': fails.append(f'table {ti} header cell {ci} bottom not 0.75pt: {bottom}')
                    elif ri==len(rows)-1:
                        if not border_is_none(top): fails.append(f'table {ti} body/last cell {ri+1},{ci} unexpected top border: {top}')
                        if bottom.get('val')!='single' or bottom.get('sz')!='12': fails.append(f'table {ti} last-row cell {ci} bottom not 1.5pt: {bottom}')
                    else:
                        if not border_is_none(top): fails.append(f'table {ti} body cell {ri+1},{ci} unexpected top border: {top}')
                        if not border_is_none(bottom): fails.append(f'table {ti} body cell {ri+1},{ci} unexpected bottom border: {bottom}')
        # Font/paragraph rules.
        for pa in tbl.findall('.//w:p',NS):
            t=text(pa)
            if not t: continue
            if jc(pa)!='center': fails.append(f'table {ti} cell not centered: {t[:40]} jc={jc(pa)}')
            sp=spacing(pa)
            if sp.get('lineRule')!='auto': fails.append(f'table {ti} cell line spacing not single/auto: {t[:40]} {sp}')
            ia=ind(pa)
            if ia.get('firstLineChars') not in (None,'0') or ia.get('firstLine') not in (None,'0'): fails.append(f'table {ti} cell has first-line indent: {t[:40]} {ia}')
            for rp in [x for x in run_props(pa) if any(x.values())]:
                if rp.get('east')!='宋体' or rp.get('ascii')!='Times New Roman' or rp.get('hAnsi')!='Times New Roman' or rp.get('size')!='21':
                    fails.append(f'table {ti} cell font wrong: {t[:40]} props={rp}')
    print('WARNING_COUNT',len(warns))
    for w in warns[:80]: print('WARN',w)
    print('FAILURE_COUNT',len(fails))
    for f in fails[:200]: print('FAIL',f)
    print('RESULT','通过' if not fails else '未通过')
    return 0 if not fails else 1
if __name__=='__main__': raise SystemExit(main(sys.argv[1]))

