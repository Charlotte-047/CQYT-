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
def jc(p):
    ppr=p.find(q('pPr')); n=None if ppr is None else ppr.find(q('jc'))
    return None if n is None else n.get(q('val'))
def ind(p):
    ppr=p.find(q('pPr')); n=None if ppr is None else ppr.find(q('ind'))
    return {} if n is None else {k.split('}')[-1]:v for k,v in n.attrib.items()}
def spacing(p):
    ppr=p.find(q('pPr')); n=None if ppr is None else ppr.find(q('spacing'))
    return {} if n is None else {k.split('}')[-1]:v for k,v in n.attrib.items()}
def has_pbb(p):
    ppr=p.find(q('pPr'))
    return ppr is not None and ppr.find(q('pageBreakBefore')) is not None
def rp(r):
    rpr=r.find(q('rPr'))
    if rpr is None: return {}
    rf=rpr.find(q('rFonts')); sz=rpr.find(q('sz')); b=rpr.find(q('b'))
    return {'east':None if rf is None else rf.get(q('eastAsia')), 'ascii':None if rf is None else rf.get(q('ascii')), 'hAnsi':None if rf is None else rf.get(q('hAnsi')), 'size':None if sz is None else sz.get(q('val')), 'bold':False if b is None else b.get(q('val')) not in ('0','false','False')}
def first_run(p):
    r=p.find('./w:r',NS)
    return {} if r is None else rp(r)
def body_like(t):
    return len(t)>90 or bool(re.search(r'[。；;]$',t))
def lvl_by_text(t):
    if body_like(t): return None
    if t in ('致谢','参考文献'): return 1
    if re.match(r'^\d+\.\d+\.\d+\s*\S.{0,80}$',t): return 3
    if re.match(r'^\d+\.\d+\s*\S.{0,75}$',t): return 2
    if re.match(r'^\d+\s+\S.{0,60}$',t): return 1
    return None
def is_toc_line(t): return bool(re.match(r'^(\d+(?:\.\d+)*\s+.+|致谢|参考文献)\s*\d+$',t))

def verify(path):
    path=Path(path); fails=[]; warns=[]; counts={1:0,2:0,3:0}; polluted=0
    with ZipFile(path) as z: root=ET.fromstring(z.read('word/document.xml'))
    body=root.find(q('body')); paras=body.findall('./w:p',NS)
    in_toc=False
    for idx,p in enumerate(paras):
        t=text(p)
        if not t: continue
        if t in ('目  录','目    录'):
            in_toc=True; continue
        if in_toc:
            if lvl_by_text(t)==1 and not is_toc_line(t): in_toc=False
            else: continue
        lvl=lvl_by_text(t); st=style(p); fr=first_run(p)
        if lvl:
            counts[lvl]+=1
            exp_style=str(2+lvl); exp_size={1:'32',2:'30',3:'28'}[lvl]; exp_align={1:'center',2:'left',3:'left'}[lvl]
            if st!=exp_style: fails.append(f'L{lvl} style wrong idx={idx} text={t} style={st} expected={exp_style}')
            if fr.get('east')!='黑体' or fr.get('ascii')!='Times New Roman' or fr.get('hAnsi')!='Times New Roman' or fr.get('size')!=exp_size or fr.get('bold'):
                fails.append(f'L{lvl} font wrong idx={idx} text={t} props={fr} expected 黑体/TNR size={exp_size} non-bold')
            if jc(p)!=exp_align: fails.append(f'L{lvl} align wrong idx={idx} text={t} jc={jc(p)} expected={exp_align}')
            ia=ind(p)
            if ia.get('firstLine') not in (None,'0') and ia.get('firstLineChars') not in (None,'0'):
                fails.append(f'L{lvl} indent wrong idx={idx} text={t} ind={ia}')
            if lvl==1 and not has_pbb(p): fails.append(f'L1 missing pageBreakBefore idx={idx} text={t}')
        else:
            # Any non-heading paragraph styled as heading/black heading size is pollution.
            if st in ('3','4','5','Heading1','Heading2','Heading3') or (fr.get('east')=='黑体' and fr.get('size') in ('32','30','28')):
                # allow known frontmatter labels/keywords to be handled elsewhere but warn.
                if t in ('摘  要','摘    要','ABSTRACT','目  录','目    录') or t.startswith(('关键词','Key words','Keywords')):
                    warns.append(f'frontmatter heading-like idx={idx} text={t} style={st} props={fr}')
                else:
                    polluted+=1; fails.append(f'non-heading polluted idx={idx} text={t[:100]} style={st} props={fr}')
    print('FILE:',path)
    print('HEADING_COUNTS:',json.dumps(counts,ensure_ascii=False))
    print('WARNING_COUNT:',len(warns))
    for w in warns[:50]: print('WARN:',w)
    print('FAILURE_COUNT:',len(fails))
    for f in fails[:160]: print('FAIL:',f)
    print('RESULT:','通过' if not fails else '未通过')
    return 0 if not fails else 1
if __name__=='__main__': raise SystemExit(verify(sys.argv[1]))
