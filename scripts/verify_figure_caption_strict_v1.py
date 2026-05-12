from __future__ import annotations
from zipfile import ZipFile
from pathlib import Path
from lxml import etree as ET
import re, sys

W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS={'w':W.strip('{}')}
def q(t): return f'{W}{t}'
def text(p): return ''.join(t.text or '' for t in p.findall('.//w:t',NS)).strip()
def has_drawing(p): return bool(p.findall('.//w:drawing',NS))
def style(p):
    ppr=p.find(q('pPr')); ps=None if ppr is None else ppr.find(q('pStyle'))
    return '' if ps is None else ps.get(q('val')) or ''
def jc(p):
    ppr=p.find(q('pPr')); n=None if ppr is None else ppr.find(q('jc'))
    return '' if n is None else (n.get(q('val')) or '')
def spacing(p):
    ppr=p.find(q('pPr')); n=None if ppr is None else ppr.find(q('spacing'))
    return {} if n is None else {k.split('}')[-1]:v for k,v in n.attrib.items()}
def ppr_flags(p):
    ppr=p.find(q('pPr'))
    return set() if ppr is None else {c.tag.split('}')[-1] for c in ppr}
def run_props(p):
    props=[]
    for r in p.findall('./w:r',NS):
        rpr=r.find(q('rPr')); rf=None if rpr is None else rpr.find(q('rFonts')); sz=None if rpr is None else rpr.find(q('sz'))
        props.append({
            'east':None if rf is None else rf.get(q('eastAsia')),
            'ascii':None if rf is None else rf.get(q('ascii')),
            'hAnsi':None if rf is None else rf.get(q('hAnsi')),
            'size':None if sz is None else sz.get(q('val')),
        })
    return props
def is_real_body_h1(p):
    return p.tag==q('p') and style(p) in ('2','Heading1') and bool(re.match(r'^\d+\s+\S+', text(p)))

def is_fig_caption(t): return bool(re.match(r'^图\s*\d+(?:\.\d+)*\s+\S+', t or ''))
def is_table_caption(t): return bool(re.match(r'^表\s*\d+(?:\.\d+)*\s+\S+', t or ''))
def is_blank(p): return text(p)=='' and not has_drawing(p)
def main(doc):
    p=Path(doc); fails=[]; warns=[]
    with ZipFile(p) as z:
        root=ET.fromstring(z.read('word/document.xml'))
    body=root.find(q('body')); children=list(body)
    start=0
    for si,n in enumerate(children):
        if is_real_body_h1(n):
            start=si; break
    pic_indices=[i for i,n in enumerate(children) if i>=start and n.tag==q('p') and has_drawing(n)]
    cap_indices=[i for i,n in enumerate(children) if i>=start and n.tag==q('p') and is_fig_caption(text(n))]
    if len(cap_indices)<len(pic_indices): fails.append(f'figure caption count less than pictures: pics={len(pic_indices)} captions={len(cap_indices)}')
    if len(cap_indices)>len(pic_indices): warns.append(f'figure caption count greater than pictures: pics={len(pic_indices)} captions={len(cap_indices)}')
    seen_caps={}
    for ci in cap_indices:
        seen_caps.setdefault(re.sub(r'\s+',' ',text(children[ci])),[]).append(ci+1)
    for cap, rows in seen_caps.items():
        if len(rows)>1: fails.append(f'duplicate figure caption: {cap} at paragraphs {rows}')
    for i in pic_indices:
        pic=children[i]
        if jc(pic)!='center': fails.append(f'picture paragraph not centered at p{i+1}: jc={jc(pic)}')
        flags=ppr_flags(pic)
        if 'keepNext' not in flags or 'keepLines' not in flags: fails.append(f'picture paragraph missing keepNext/keepLines at p{i+1}: {flags}')
        j=i+1
        blanks_between=0
        while j<len(children) and children[j].tag==q('p') and is_blank(children[j]):
            blanks_between+=1; j+=1
        if blanks_between: fails.append(f'blank paragraph between picture and caption at p{i+1}: blanks={blanks_between}')
        if j>=len(children) or children[j].tag!=q('p') or not is_fig_caption(text(children[j])):
            near=[text(children[k]) for k in range(i+1,min(len(children),i+5)) if children[k].tag==q('p')]
            fails.append(f'picture not immediately followed by figure caption at p{i+1}; next={near}')
            continue
        cap=children[j]; t=text(cap)
        if jc(cap)!='center': fails.append(f'figure caption not centered at p{j+1}: {t} jc={jc(cap)}')
        sp=spacing(cap)
        if sp.get('line')!='440' or sp.get('lineRule')!='exact': fails.append(f'figure caption line spacing not fixed 22pt at p{j+1}: {sp}')
        flags=ppr_flags(cap)
        if 'keepNext' not in flags or 'keepLines' not in flags: fails.append(f'figure caption missing keepNext/keepLines at p{j+1}: {flags}')
        rps=[rp for rp in run_props(cap) if any(rp.values())]
        if not rps: fails.append(f'figure caption missing run props at p{j+1}: {t}')
        for rp in rps:
            if rp.get('east')!='宋体' or rp.get('size')!='21': fails.append(f'figure caption font wrong at p{j+1}: {t} props={rp}')
            if rp.get('ascii')!='Times New Roman' or rp.get('hAnsi')!='Times New Roman': fails.append(f'figure caption latin font wrong at p{j+1}: {t} props={rp}')
        # Rules: one blank before picture and one after caption when body paragraphs exist around it.
        if i>0 and children[i-1].tag==q('p') and text(children[i-1]) and not is_fig_caption(text(children[i-1])) and not is_table_caption(text(children[i-1])):
            warns.append(f'no blank line before picture p{i+1}')
        if j+1<len(children) and children[j+1].tag==q('p') and text(children[j+1]) and not has_drawing(children[j+1]):
            warns.append(f'no blank line after figure caption p{j+1}')
    print('FIGURE_COUNT',len(pic_indices))
    print('FIGURE_CAPTION_COUNT',len(cap_indices))
    print('WARNING_COUNT',len(warns))
    for w in warns[:80]: print('WARN',w)
    print('FAILURE_COUNT',len(fails))
    for f in fails[:120]: print('FAIL',f)
    print('RESULT','通过' if not fails else '未通过')
    return 0 if not fails else 1
if __name__=='__main__': raise SystemExit(main(sys.argv[1]))
