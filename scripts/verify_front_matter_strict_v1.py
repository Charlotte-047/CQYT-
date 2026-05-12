from __future__ import annotations
from zipfile import ZipFile
from pathlib import Path
from lxml import etree as ET
import re, sys

W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS={'w':W.strip('{}')}
def q(t): return f'{W}{t}'
def text(p): return ''.join(t.text or '' for t in p.findall('.//w:t',NS)).strip()
def raw(p): return ''.join(t.text or '' for t in p.findall('.//w:t',NS))
def norm(t): return re.sub(r'\s+','',t or '')
def jc(p):
    ppr=p.find(q('pPr')); n=None if ppr is None else ppr.find(q('jc'))
    return '' if n is None else (n.get(q('val')) or '')
def ind(p):
    ppr=p.find(q('pPr')); n=None if ppr is None else ppr.find(q('ind'))
    return {} if n is None else {k.split('}')[-1]:v for k,v in n.attrib.items()}
def spacing(p):
    ppr=p.find(q('pPr')); n=None if ppr is None else ppr.find(q('spacing'))
    return {} if n is None else {k.split('}')[-1]:v for k,v in n.attrib.items()}
def run_props(p):
    vals=[]
    for r in p.findall('./w:r',NS):
        rpr=r.find(q('rPr')); rf=None if rpr is None else rpr.find(q('rFonts')); sz=None if rpr is None else rpr.find(q('sz')); b=None if rpr is None else rpr.find(q('b'))
        vals.append({'east':None if rf is None else rf.get(q('eastAsia')),'ascii':None if rf is None else rf.get(q('ascii')),'hAnsi':None if rf is None else rf.get(q('hAnsi')),'size':None if sz is None else sz.get(q('val')),'bold':(b is not None and b.get(q('val')) not in ('0','false','False'))})
    return vals
def has_page_break(p): return bool(p.findall('.//w:br[@w:type="page"]',NS))
def page_break_before(p):
    ppr=p.find(q('pPr')); return ppr is not None and ppr.find(q('pageBreakBefore')) is not None
def has_font(p,east=None,ascii=None,size=None,bold=None):
    rps=[x for x in run_props(p) if any(x.values())]
    if not rps: return False
    for rp in rps:
        if east is not None and rp.get('east')!=east: return False
        if ascii is not None and (rp.get('ascii')!=ascii or rp.get('hAnsi')!=ascii): return False
        if size is not None and rp.get('size')!=size: return False
        if bold is not None and bool(rp.get('bold'))!=bold: return False
    return True
def is_cn_abs(t): return norm(t)=='摘要'
def is_toc(t): return norm(t)=='目录'
def find_idx(ps,pred):
    for i,p in enumerate(ps):
        if pred(text(p)): return i
    return None
def para_between(ps,a,b): return ps[a+1:b] if a is not None and b is not None and a<b else []
def body_para_ok(p,font='宋体'):
    sp=spacing(p); ia=ind(p)
    return jc(p)=='both' and ia.get('firstLineChars')=='200' and sp.get('line')=='440' and sp.get('lineRule')=='exact'
def load_paras(doc):
    with ZipFile(Path(doc)) as z:
        root=ET.fromstring(z.read('word/document.xml'))
    return root.findall('.//w:body/w:p',NS)

def main(doc, src=None):
    fails=[]; warns=[]; p=Path(doc)
    ps=load_paras(p)
    src_ps=load_paras(src) if src else ps
    src_has_stmt=find_idx(src_ps, lambda t: norm(t)=='原创性声明') is not None
    src_has_cn=find_idx(src_ps, is_cn_abs) is not None
    src_has_en=find_idx(src_ps, lambda t: norm(t).upper()=='ABSTRACT') is not None
    idx_stmt=find_idx(ps, lambda t: norm(t)=='原创性声明')
    idx_cn=find_idx(ps, is_cn_abs)
    idx_en=find_idx(ps, lambda t: norm(t).upper()=='ABSTRACT')
    idx_toc=find_idx(ps, is_toc)
    idx_body=find_idx(ps, lambda t: bool(re.match(r'^1\s+\S+',t or '')))
    required=[('原创性声明',idx_stmt),('摘    要',idx_cn),('ABSTRACT',idx_en),('目录',idx_toc),('正文1级标题',idx_body)]
    required_present=[('原创性声明',idx_stmt,src_has_stmt),('摘    要',idx_cn,src_has_cn),('ABSTRACT',idx_en,src_has_en),('目录',idx_toc,True),('正文1级标题',idx_body,True)]
    for name,idx,needed in required_present:
        if idx is None and needed: fails.append(f'front matter marker missing: {name}')
        elif idx is None and not needed: warns.append(f'front matter absent in source, skipped: {name}')
    if idx_stmt is not None and src_has_stmt:
        p0=ps[idx_stmt]
        if jc(p0)!='center' or not has_font(p0,east='宋体',ascii='Times New Roman',size='44',bold=True): fails.append(f'originality title format wrong: jc={jc(p0)} props={run_props(p0)}')
        for k,p1 in enumerate(para_between(ps,idx_stmt,idx_cn if idx_cn is not None else min(len(ps),idx_stmt+20)),idx_stmt+2):
            t=text(p1)
            if not t: continue
            if re.search(r'(本人签名|签名|姓名|日期|年\s*月\s*日)',t):
                if jc(p1)!='right': fails.append(f'originality signature/date not right aligned p{k}: {t} jc={jc(p1)}')
                ia=ind(p1)
                if ia.get('firstLineChars') not in (None,'0') or ia.get('firstLine') not in (None,'0'): fails.append(f'originality signature/date has indent p{k}: {t} {ia}')
            else:
                if not body_para_ok(p1): fails.append(f'originality body paragraph format wrong p{k}: {t[:40]} jc={jc(p1)} ind={ind(p1)} sp={spacing(p1)}')
    if idx_cn is not None and src_has_cn:
        p0=ps[idx_cn]
        if text(p0)!='摘    要': fails.append(f'Chinese abstract title text should be spaced 摘    要: {text(p0)!r}')
        if jc(p0)!='center' or not has_font(p0,east='黑体',size='32'): fails.append(f'Chinese abstract title format wrong: jc={jc(p0)} props={run_props(p0)}')
        saw_kw=False
        for k,p1 in enumerate(para_between(ps,idx_cn,idx_en if idx_en is not None else min(len(ps),idx_cn+30)),idx_cn+2):
            t=text(p1)
            if not t: continue
            if t.startswith('关键词'):
                saw_kw=True
                if jc(p1) not in ('left','both',''): fails.append(f'Chinese keywords not left/top aligned p{k}: jc={jc(p1)}')
                if '；' not in t and ';' in t: fails.append(f'Chinese keywords should use Chinese semicolon p{k}: {t}')
                if re.search(r'[。；;,.，]$',t): fails.append(f'Chinese keywords should not end with punctuation p{k}: {t}')
            else:
                if not body_para_ok(p1): fails.append(f'Chinese abstract body format wrong p{k}: {t[:40]} jc={jc(p1)} ind={ind(p1)} sp={spacing(p1)}')
        if not saw_kw: fails.append('Chinese keywords line missing')
    if idx_en is not None and src_has_en:
        p0=ps[idx_en]
        if text(p0)!='ABSTRACT': fails.append(f'English abstract title must be ABSTRACT: {text(p0)!r}')
        if jc(p0)!='center' or not has_font(p0,east='Times New Roman',ascii='Times New Roman',size='32',bold=True): fails.append(f'English abstract title format wrong: jc={jc(p0)} props={run_props(p0)}')
        if not (has_page_break(p0) or page_break_before(p0)): warns.append('English abstract title does not show explicit page break before')
        saw_kw=False
        for k,p1 in enumerate(para_between(ps,idx_en,idx_toc if idx_toc is not None else min(len(ps),idx_en+30)),idx_en+2):
            t=text(p1)
            if not t: continue
            if re.match(r'^(Key words|Keywords)\b',t):
                saw_kw=True
                if jc(p1) not in ('left','both',''): fails.append(f'English keywords not left/top aligned p{k}: jc={jc(p1)}')
                if '；' in t: fails.append(f'English keywords should use English semicolon p{k}: {t}')
                if re.search(r'[。；;,.，]$',t): fails.append(f'English keywords should not end with punctuation p{k}: {t}')
                if '; ' not in t and ';' in t: fails.append(f'English keywords semicolon should be followed by one space p{k}: {t}')
            else:
                if not body_para_ok(p1): fails.append(f'English abstract body format wrong p{k}: {t[:40]} jc={jc(p1)} ind={ind(p1)} sp={spacing(p1)}')
                if not has_font(p1,east='Times New Roman',ascii='Times New Roman',size='24',bold=False): fails.append(f'English abstract body font wrong p{k}: {t[:40]} props={run_props(p1)}')
        if not saw_kw: fails.append('English keywords line missing')
    if idx_toc is not None:
        p0=ps[idx_toc]
        if jc(p0)!='center' or not has_font(p0,east='宋体',size='32',bold=False): fails.append(f'TOC title format wrong: jc={jc(p0)} props={run_props(p0)}')
    # Structure order.
    order=[idx for _,idx in required if idx is not None]
    if order != sorted(order): fails.append(f'front matter order wrong: {required}')
    print('FRONT_MARKERS', {name:idx for name,idx in required})
    print('WARNING_COUNT',len(warns))
    for w in warns[:80]: print('WARN',w)
    print('FAILURE_COUNT',len(fails))
    for f in fails[:160]: print('FAIL',f)
    print('RESULT','通过' if not fails else '未通过')
    return 0 if not fails else 1
if __name__=='__main__': raise SystemExit(main(sys.argv[-1], sys.argv[1] if len(sys.argv)>2 else None))


