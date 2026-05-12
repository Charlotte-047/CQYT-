from zipfile import ZipFile
from pathlib import Path
from lxml import etree as ET
import sys,re,json,hashlib

W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
WP='{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
NS={'w':W.strip('{}'),'wp':WP.strip('{}')}
FONT_HEI=chr(0x9ed1)+chr(0x4f53)
FONT_SONG=chr(0x5b8b)+chr(0x4f53)
TXT_TOC=chr(0x76ee)+chr(0x5f55)
TXT_REFS=chr(0x53c2)+chr(0x8003)+chr(0x6587)+chr(0x732e)
def q(t): return f'{W}{t}'
def text(n): return ''.join(t.text or '' for t in n.findall('.//w:t',NS)).strip()
def raw(n): return ''.join(t.text or '' for t in n.findall('.//w:t',NS))
def body(path):
    with ZipFile(path) as z:
        root=ET.fromstring(z.read('word/document.xml'))
    return root.find(q('body'))
def paras(body): return body.findall('./w:p',NS)
def tables(body): return body.findall('.//w:tbl',NS)
def drawings(body): return body.findall('.//w:drawing',NS)
def pstyle(p):
    ppr=p.find(q('pPr')); ps=None if ppr is None else ppr.find(q('pStyle'))
    return '' if ps is None else (ps.get(q('val')) or '')
def jc(p):
    ppr=p.find(q('pPr')); j=None if ppr is None else ppr.find(q('jc'))
    return None if j is None else j.get(q('val'))
def ind(p):
    ppr=p.find(q('pPr')); n=None if ppr is None else ppr.find(q('ind'))
    return {} if n is None else {k.split('}')[-1]:v for k,v in n.attrib.items()}
def spacing(p):
    ppr=p.find(q('pPr')); n=None if ppr is None else ppr.find(q('spacing'))
    return {} if n is None else {k.split('}')[-1]:v for k,v in n.attrib.items()}
def run_props(p):
    r=p.find('./w:r',NS); rpr=None if r is None else r.find(q('rPr')); rf=None if rpr is None else rpr.find(q('rFonts')); sz=None if rpr is None else rpr.find(q('sz'))
    return {} if r is None else {'east':None if rf is None else rf.get(q('eastAsia')),'ascii':None if rf is None else rf.get(q('ascii')),'hAnsi':None if rf is None else rf.get(q('hAnsi')),'size':None if sz is None else sz.get(q('val'))}
def body_like(t): return len(t)>70 or bool(re.search(r'[。；;]$',t))
def hlevel(t):
    if body_like(t): return None
    if t in ('致谢','参考文献'): return 1
    if re.match(r'^\d+\.\d+\.\d+\s*\S.{0,70}$',t): return 3
    if re.match(r'^\d+\.\d+\s*\S.{0,65}$',t): return 2
    if re.match(r'^\d+\s+\S.{0,55}$',t): return 1
    return None
def norm(t): return re.sub(r'\s+','',t or '')
def is_toc_title(t): return norm(t)=='目录'
def is_toc_line(t): return bool(re.match(r'^(\d+(?:\.\d+)*\s+.+|致谢|参考文献)\s*\d+$',t))
def sha_text(items): return hashlib.sha256('\n'.join(items).encode('utf-8','ignore')).hexdigest()[:16]
def collect_texts(ps): return [text(p) for p in ps if text(p)]
def collect_headings(ps):
    out=[]; in_toc=False
    for p in ps:
        t=text(p)
        if is_toc_title(t): in_toc=True; continue
        if in_toc:
            if is_toc_line(t): continue
            if hlevel(t)==1: in_toc=False
            else: continue
        lv=hlevel(t)
        if lv: out.append((lv,t))
    return out
def table_texts(tbls):
    res=[]
    for tbl in tbls:
        cells=[]
        for tc in tbl.findall('.//w:tc',NS):
            ct=' '.join(text(p) for p in tc.findall('.//w:p',NS) if text(p))
            if ct: cells.append(ct)
        res.append('|'.join(cells))
    return res

def audit(src,out):
    bs,bo=body(src),body(out)
    ps,po=paras(bs),paras(bo)
    src_texts=collect_texts(ps); out_texts=collect_texts(po)
    src_heads=collect_headings(ps); out_heads=collect_headings(po)
    src_tbls, out_tbls = table_texts(tables(bs)), table_texts(tables(bo))
    fails=[]; warns=[]; metrics={}
    metrics.update({
        'src_paragraphs':len(ps),'out_paragraphs':len(po),
        'src_nonempty_paragraphs':len(src_texts),'out_nonempty_paragraphs':len(out_texts),
        'src_headings':len(src_heads),'out_headings':len(out_heads),
        'src_tables':len(src_tbls),'out_tables':len(out_tbls),
        'src_drawings':len(drawings(bs)),'out_drawings':len(drawings(bo)),
        'src_text_hash':sha_text(src_texts),'out_text_hash':sha_text(out_texts),
        'src_table_hash':sha_text(src_tbls),'out_table_hash':sha_text(out_tbls),
    })
    if len(src_heads)!=len(out_heads): warns.append(f'heading count changed: {len(src_heads)} -> {len(out_heads)}')
    missing=[h for h in src_heads if h not in out_heads]
    extra=[h for h in out_heads if h not in src_heads]
    if missing: warns.append(f'missing headings: {missing[:10]}')
    if extra: warns.append(f'extra headings: {extra[:10]}')
    if len(src_tbls)!=len(out_tbls): fails.append(f'table count changed: {len(src_tbls)} -> {len(out_tbls)}')
    if len(drawings(bo)) < len(drawings(bs)): fails.append(f'drawings lost: {len(drawings(bs))} -> {len(drawings(bo))}')
    # Text preservation heuristic: every long source paragraph should appear in output after whitespace normalization.
    out_join=norm('\n'.join(out_texts))
    lost=[]
    for t in src_texts:
        nt=norm(t)
        if len(nt)>=40 and nt not in out_join:
            lost.append(t[:100])
    if lost: fails.append(f'long paragraphs not found in output: {len(lost)} examples={lost[:8]}')
    # Formatting coverage on output.
    toc_count=sum(1 for p in po if is_toc_line(text(p)))
    if not any(is_toc_title(text(p)) for p in po): fails.append('TOC title missing')
    if toc_count < max(3,len(out_heads)//2): fails.append(f'TOC entries suspiciously few: {toc_count}')
    bad_body=[]; bad_heads=[]; in_toc=False; in_body=False; in_refs=False
    for i,p in enumerate(po):
        t=text(p)
        if is_toc_title(t): in_toc=True; continue
        if in_toc:
            if is_toc_line(t): continue
            if hlevel(t)==1: in_toc=False
            else: continue
        lv=hlevel(t); rp=run_props(p)
        if norm(t)==TXT_REFS:
            in_refs=True
        if lv:
            exp_size={1:'32',2:'30',3:'28'}[lv]
            exp_style=str(1+lv)
            if pstyle(p)!=exp_style or rp.get('east')!=FONT_HEI or rp.get('size')!=exp_size:
                bad_heads.append((i,t,pstyle(p),rp))
            if lv==1: in_body=True
        elif in_refs:
            continue
        elif in_body and t and not t.startswith((chr(0x5173)+chr(0x952e)+chr(0x8bcd),'Key words','Keywords')) and not re.match(r'^['+chr(0x56fe)+chr(0x8868)+r']\d+',t):
            if body_like(t):
                ia=ind(p)
                if rp.get('east')!=FONT_SONG or rp.get('ascii')!='Times New Roman' or rp.get('size')!='24' or jc(p)!='both' or ia.get('firstLineChars')!='200':
                    bad_body.append((i,t[:60],jc(p),ia,rp))
    if bad_heads: fails.append(f'bad heading formats: {len(bad_heads)} examples={bad_heads[:8]}')
    if bad_body: fails.append(f'bad body paragraph formats: {len(bad_body)} examples={bad_body[:8]}')
    # Table formatting sample.
    bad_tables=[]
    for ti,tbl in enumerate(tables(bo),1):
        for p in tbl.findall('.//w:p',NS):
            if text(p):
                rp=run_props(p); sp=spacing(p)
                if jc(p)!='center' or sp.get('lineRule')!='auto' or rp.get('east')!='宋体' or rp.get('size')!='21':
                    bad_tables.append((ti,text(p)[:40],jc(p),sp,rp)); break
    if bad_tables: fails.append(f'bad table formatting: {len(bad_tables)} examples={bad_tables[:8]}')
    print('FORMAT_COVERAGE_METRICS', json.dumps(metrics,ensure_ascii=False))
    print('WARNING_COUNT',len(warns))
    for w in warns: print('WARN',w)
    print('FAILURE_COUNT',len(fails))
    for f in fails: print('FAIL',f)
    print('RESULT', '通过' if not fails else '未通过')
    return 0 if not fails else 1

if __name__=='__main__':
    raise SystemExit(audit(Path(sys.argv[1]), Path(sys.argv[2])))
