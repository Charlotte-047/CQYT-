from zipfile import ZipFile
from pathlib import Path
from lxml import etree as ET
import sys,re,json

W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS={'w':W.strip('{}')}
def q(t): return f'{W}{t}'
def text(p): return ''.join(t.text or '' for t in p.findall('.//w:t',NS)).strip()
def style(p):
    ppr=p.find(q('pPr')); ps=None if ppr is None else ppr.find(q('pStyle'))
    return '' if ps is None else (ps.get(q('val')) or '')
def rp(p):
    r=p.find('./w:r',NS)
    if r is None: return {}
    rpr=r.find(q('rPr'))
    if rpr is None: return {}
    rf=rpr.find(q('rFonts')); sz=rpr.find(q('sz')); b=rpr.find(q('b'))
    return {'east':None if rf is None else rf.get(q('eastAsia')), 'ascii':None if rf is None else rf.get(q('ascii')), 'hAnsi':None if rf is None else rf.get(q('hAnsi')), 'size':None if sz is None else sz.get(q('val')), 'bold':False if b is None else b.get(q('val')) not in ('0','false','False')}
def body_like(t): return len(t)>90 or bool(re.search(r'[。；;]$',t))
def lvl(t):
    if body_like(t): return None
    if t in ('致谢','参考文献'): return 1
    if re.match(r'^\d+\.\d+\.\d+\s*\S.{0,80}$',t): return 3
    if re.match(r'^\d+\.\d+\s*\S.{0,75}$',t): return 2
    if re.match(r'^\d+\s+\S.{0,60}$',t): return 1
    return None
def norm(t): return re.sub(r'\s+',' ',t).strip()
def is_toc_line(t): return bool(re.match(r'^(\d+(?:\.\d+)*\s+.+|致谢|参考文献)\s*\d+$',t))
def load_paras(path):
    with ZipFile(path) as z: root=ET.fromstring(z.read('word/document.xml'))
    return root.find(q('body')).findall('./w:p',NS)
def extract_expected_from_source(src):
    out=[]; in_toc=False; body_started=False
    for p in load_paras(src):
        t=text(p)
        if not t: continue
        if t in ('目  录','目    录'):
            in_toc=True; continue
        if in_toc:
            if is_toc_line(t): continue
            if lvl(t)==1: in_toc=False
            else: continue
        l=lvl(t)
        if l:
            body_started=True; out.append((l,norm(t)))
    # If source headings are too polluted, prefer TOC-derived expected list.
    if len(out) < 5:
        out=[]; in_toc=False
        for p in load_paras(src):
            t=text(p)
            if t in ('目  录','目    录'): in_toc=True; continue
            if in_toc:
                if is_toc_line(t):
                    tt=re.sub(r'\s*\d+$','',t)
                    l=lvl(tt)
                    if l: out.append((l,norm(tt)))
                elif lvl(t)==1: break
    return out
def map_output(outdoc):
    d={}; seq=[]; in_toc=False
    for p in load_paras(outdoc):
        t=text(p)
        if not t: continue
        if t in ('目  录','目    录'):
            in_toc=True; continue
        if in_toc:
            if lvl(t)==1 and not is_toc_line(t): in_toc=False
            else: continue
        l=lvl(t)
        if l:
            key=norm(t); d[key]=(l,style(p),rp(p)); seq.append((l,key,style(p),rp(p)))
    return d,seq
def verify(src,outdoc):
    fails=[]; expected=extract_expected_from_source(src); got,seq=map_output(outdoc)
    for l,t in expected:
        if t not in got:
            fails.append(f'missing expected heading L{l}: {t}')
            continue
        gl,st,props=got[t]
        exp_style=str(1+l); exp_size={1:'32',2:'30',3:'28'}[l]
        if gl!=l: fails.append(f'level mismatch for {t}: expected L{l}, got L{gl}')
        if st!=exp_style: fails.append(f'style mismatch for {t}: expected {exp_style}, got {st}')
        if props.get('east')!='黑体' or props.get('size')!=exp_size:
            fails.append(f'font mismatch for {t}: expected 黑体 size {exp_size}, got {props}')
    exp_set={t for _,t in expected}
    for gl,t,st,props in seq:
        if t not in exp_set and not t.startswith(('摘','ABSTRACT','目')):
            fails.append(f'extra output heading not in source expected L{gl}: {t}')
    print('SOURCE_EXPECTED_COUNT',len(expected))
    for x in expected: print('EXPECT',x)
    print('OUTPUT_HEADING_COUNT',len(seq))
    print('FAILURE_COUNT',len(fails))
    for f in fails[:160]: print('FAIL',f)
    print('RESULT','通过' if not fails else '未通过')
    return 0 if not fails else 1
if __name__=='__main__': raise SystemExit(verify(sys.argv[1],sys.argv[2]))
