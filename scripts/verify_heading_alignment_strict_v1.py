from zipfile import ZipFile
from pathlib import Path
from lxml import etree as ET
import re, sys
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS={'w':W.strip('{}')}
def q(t): return f'{W}{t}'
def text(p): return ''.join(t.text or '' for t in p.findall('.//w:t',NS)).strip()
def style(p):
    ppr=p.find(q('pPr')); ps=None if ppr is None else ppr.find(q('pStyle'))
    return '' if ps is None else ps.get(q('val')) or ''
def jc(p):
    ppr=p.find(q('pPr')); j=None if ppr is None else ppr.find(q('jc'))
    return '' if j is None else (j.get(q('val')) or '')
def has_list_format(p):
    ppr=p.find(q('pPr'))
    return ppr is not None and (ppr.find(q('numPr')) is not None or ppr.find(q('tabs')) is not None)

def is_real_body_start(p):
    t=text(p)
    return style(p) in ('2','Heading1') and bool(re.match(r'^\d+\s+\S',t))

def hlevel(t):
    if len(t)>90 or re.search(r'[。；;]$',t): return None
    if t in ('致谢','参考文献'): return 1
    if re.match(r'^\d+\.\d+\.\d+\s*\S.{0,70}$',t): return 3
    if re.match(r'^\d+\.\d+\s*\S.{0,65}$',t): return 2
    if re.match(r'^\d+\s+\S.{0,55}$',t): return 1
    return None
def main(doc):
    root=ET.fromstring(ZipFile(doc).read('word/document.xml'))
    fails=[]
    paras=root.findall('.//w:body/w:p',NS)
    start=0
    for si,sp in enumerate(paras):
        if is_real_body_start(sp):
            start=si; break
    for i,p in enumerate(paras[start:],start+1):
        t=text(p); lvl=hlevel(t)
        if lvl in (1,2,3) and has_list_format(p): fails.append(f'heading has list/tabs p{i}: {t} style={style(p)}')
        if lvl==1:
            if jc(p)!='center': fails.append(f'H1 not center p{i}: {t} jc={jc(p)} style={style(p)}')
            if style(p) not in ('2','Heading1'): fails.append(f'H1 wrong style p{i}: {t} style={style(p)}')
        elif lvl==2:
            if jc(p)!='left': fails.append(f'H2 not left p{i}: {t} jc={jc(p)} style={style(p)}')
            if style(p) not in ('3','Heading2'): fails.append(f'H2 wrong style p{i}: {t} style={style(p)}')
        elif lvl==3:
            if jc(p)!='left': fails.append(f'H3 not left p{i}: {t} jc={jc(p)} style={style(p)}')
            if style(p) not in ('4','Heading3'): fails.append(f'H3 wrong style p{i}: {t} style={style(p)}')
    print('HEADING_ALIGNMENT_FAILURE_COUNT',len(fails))
    for f in fails: print('FAIL',f)
    print('RESULT','通过' if not fails else '未通过')
    return 0 if not fails else 1
if __name__=='__main__': raise SystemExit(main(sys.argv[1]))
