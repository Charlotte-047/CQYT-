from zipfile import ZipFile
from pathlib import Path
from lxml import etree as ET
import sys,re,json

W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
WP='{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
NS={'w':W.strip('{}'),'wp':WP.strip('{}')}
def q(t): return f'{W}{t}'
def text(n): return ''.join(t.text or '' for t in n.findall('.//w:t',NS)).strip()
def raw(n): return ''.join(t.text or '' for t in n.findall('.//w:t',NS))
def style(p):
    ppr=p.find(q('pPr')); ps=None if ppr is None else ppr.find(q('pStyle'))
    return '' if ps is None else (ps.get(q('val')) or '')
def ppr(p): return p.find(q('pPr'))
def child_val(p,tag):
    pp=ppr(p); n=None if pp is None else pp.find(q(tag))
    return None if n is None else n.get(q('val'))
def jc(p): return child_val(p,'jc')
def spacing(p):
    pp=ppr(p); n=None if pp is None else pp.find(q('spacing'))
    return {} if n is None else {k.split('}')[-1]:v for k,v in n.attrib.items()}
def ind(p):
    pp=ppr(p); n=None if pp is None else pp.find(q('ind'))
    return {} if n is None else {k.split('}')[-1]:v for k,v in n.attrib.items()}
def has_pbb(p):
    pp=ppr(p); return pp is not None and pp.find(q('pageBreakBefore')) is not None
def has_br_page(p): return bool(p.findall('.//w:br[@w:type="page"]',NS))
def has_drawing(p): return bool(p.findall('.//w:drawing',NS))
def has_anchor(p): return bool(p.findall('.//wp:anchor',NS))
def has_inline(p): return bool(p.findall('.//wp:inline',NS))
def is_blank(p): return p.tag==q('p') and text(p)=='' and not has_drawing(p)
def run_props(r):
    rpr=r.find(q('rPr'))
    rf=None if rpr is None else rpr.find(q('rFonts'))
    sz=None if rpr is None else rpr.find(q('sz'))
    b=None if rpr is None else rpr.find(q('b'))
    color=None if rpr is None else rpr.find(q('color'))
    return {'east':None if rf is None else rf.get(q('eastAsia')),'ascii':None if rf is None else rf.get(q('ascii')),'hAnsi':None if rf is None else rf.get(q('hAnsi')),'size':None if sz is None else sz.get(q('val')),'bold':False if b is None else b.get(q('val')) not in ('0','false','False'),'color':None if color is None else color.get(q('val'))}
def first_run(p):
    r=p.find('./w:r',NS); return {} if r is None else run_props(r)
def body_like(t): return len(t)>90 or bool(re.search(r'[。；;]$',t))
def h_level(t):
    if body_like(t): return None
    if t in ('\u81f4\u8c22','\u53c2\u8003\u6587\u732e'): return 1
    if re.match(r'^\d+\.\d+\.\d+\s*\S.{0,80}$',t): return 3
    if re.match(r'^\d+\.\d+\s*\S.{0,75}$',t): return 2
    if re.match(r'^\d+\s+\S.{0,60}$',t): return 1
    return None
def norm_space(t): return re.sub(r'\s+','',t or '')
def is_statement_title(t): return '\u539f\u521b\u6027\u58f0\u660e' in norm_space(t) or '\u539f\u521b\u6027' in norm_space(t)
def is_cn_abs_title(t):
    n=norm_space(t)
    return n in ('\u6458\u8981','\u4e2d\u6587\u6458\u8981') or ('\u6458' in n and '\u8981' in n and len(n)<=6)
def is_en_abs_title(t): return norm_space(t).upper()=='ABSTRACT'
def is_toc_title(t): return norm_space(t)=='\u76ee\u5f55'
def is_refs_title(t): return norm_space(t)=='\u53c2\u8003\u6587\u732e'
def is_toc_line(t): return bool(re.match(r'^(\d+(?:\.\d+)*\s+.+|\u81f4\u8c22|\u53c2\u8003\u6587\u732e)\s*\d+$',t))
def is_caption(t): return bool(re.match(r'^[\u56fe\u8868]\d+[\.\-]\d+\s*\S+',t)) or t.startswith('\u6ce8\uff1a')
def is_ref(t): return bool(re.match(r'^\[\d+\]\s+',t))
def fail(fails,mod,msg): fails.append({'module':mod,'msg':msg})
def warn(warns,mod,msg): warns.append({'module':mod,'msg':msg})
def ok_body_font(fr): return fr and fr.get('east')=='\u5b8b\u4f53' and fr.get('ascii')=='Times New Roman' and fr.get('size')=='24'
def ok_table_font(rp): return rp and rp.get('east')=='\u5b8b\u4f53' and rp.get('size')=='21'

def load(path):
    with ZipFile(path) as z:
        root=ET.fromstring(z.read('word/document.xml'))
        styles=ET.fromstring(z.read('word/styles.xml')) if 'word/styles.xml' in z.namelist() else None
    return root,styles

def audit(path):
    path=Path(path); root,styles=load(path); body=root.find(q('body')); nodes=list(body); paras=[n for n in nodes if n.tag==q('p')]
    fails=[]; warns=[]; metrics={}

    markers={'statement':False,'abstract_cn':False,'abstract_en':False,'toc':False,'body':False,'refs':False}
    in_toc=False; in_body=False
    toc_rows=0; heading_counts={1:0,2:0,3:0}; body_para=0; blank_runs=[]; cur_blank=0
    for idx,p in enumerate(paras):
        t=text(p)
        if p.getparent() is body and is_blank(p): cur_blank+=1
        else:
            if cur_blank: blank_runs.append(cur_blank)
            cur_blank=0
        if is_statement_title(t): markers['statement']=True
        if is_cn_abs_title(t): markers['abstract_cn']=True
        if is_en_abs_title(t): markers['abstract_en']=True
        if is_toc_title(t): markers['toc']=True; in_toc=True; continue
        if in_toc:
            if is_toc_line(t): toc_rows+=1; continue
            if h_level(t)==1: in_toc=False
            elif t: continue
        lvl=h_level(t)
        if lvl==1 and not (is_cn_abs_title(t) or is_en_abs_title(t) or is_toc_title(t)):
            markers['body']=True; in_body=True
        if is_refs_title(t): markers['refs']=True; in_body=False
        if in_body and t and not lvl and not is_caption(t) and not is_ref(t) and not has_drawing(p): body_para+=1
        if lvl: heading_counts[lvl]+=1
    if cur_blank: blank_runs.append(cur_blank)
    metrics.update({'toc_rows':toc_rows,'heading_counts':heading_counts,'body_paragraphs':body_para,'max_consecutive_blank_paras':max(blank_runs or [0])})
    for k,v in markers.items():
        if not v:
            if k in ('toc','body'): fail(fails,'structure',f'missing required section marker: {k}')
            else: warn(warns,'structure',f'section marker not detected: {k}')
    if markers['toc'] and toc_rows<3: fail(fails,'toc','TOC detected but too few entries')
    if max(blank_runs or [0])>=4: warn(warns,'blank_pages',f'has {max(blank_runs)} consecutive blank paragraphs; possible blank page risk')

    for p in paras:
        t=text(p); fr=first_run(p)
        if is_cn_abs_title(t):
            if jc(p)!='center' or fr.get('east')!='\u9ed1\u4f53' or fr.get('size')!='32': fail(fails,'abstract',f'CN abstract title format wrong props={fr} jc={jc(p)}')
        if is_en_abs_title(t):
            if jc(p)!='center' or fr.get('ascii')!='Times New Roman' or fr.get('size')!='32': fail(fails,'abstract',f'EN abstract title format wrong props={fr} jc={jc(p)}')

    in_toc=False
    for idx,p in enumerate(paras):
        t=text(p)
        if is_toc_title(t): in_toc=True; continue
        if not in_toc: continue
        if not t: continue
        if is_toc_line(t):
            if style(p) in ('3','4','5','Heading1','Heading2','Heading3'): fail(fails,'toc',f'TOC line has heading style idx={idx}: {t}')
            # Preserved source TOC may contain page/section marks from Word layout; keep it intact.
            fr=first_run(p)
            if fr.get('east')=='\u9ed1\u4f53' and fr.get('size') in ('32','30','28'): fail(fails,'toc',f'TOC line polluted by heading font idx={idx}: {t}')
            continue
        if h_level(t)==1: break

    in_toc=False
    for idx,p in enumerate(paras):
        t=text(p)
        if not t: continue
        if is_toc_title(t): in_toc=True; continue
        if in_toc:
            if is_toc_line(t): continue
            if h_level(t)==1: in_toc=False
            else: continue
        lvl=h_level(t); st=style(p); fr=first_run(p)
        if lvl:
            exp_style=str(2+lvl); exp_size={1:'32',2:'30',3:'28'}[lvl]; exp_jc={1:'center',2:'left',3:'left'}[lvl]
            if st!=exp_style: fail(fails,'headings',f'L{lvl} style wrong idx={idx}: {t} style={st}')
            if fr.get('east')!='\u9ed1\u4f53' or fr.get('size')!=exp_size or fr.get('bold'): fail(fails,'headings',f'L{lvl} font wrong idx={idx}: {t} props={fr}')
            if jc(p)!=exp_jc: fail(fails,'headings',f'L{lvl} align wrong idx={idx}: {t} jc={jc(p)}')
            if lvl==1 and not has_pbb(p): fail(fails,'headings',f'L1 missing pageBreakBefore idx={idx}: {t}')
        elif body_like(t):
            if st in ('3','4','5','Heading1','Heading2','Heading3') or (fr.get('east')=='\u9ed1\u4f53' and fr.get('size') in ('32','30','28')):
                fail(fails,'body',f'body paragraph polluted as heading idx={idx}: {t[:80]}')

    in_body=False; in_toc=False
    for idx,p in enumerate(paras):
        t=text(p); lvl=h_level(t)
        if is_toc_title(t): in_toc=True; continue
        if in_toc:
            if is_toc_line(t): continue
            if lvl==1: in_toc=False
            else: continue
        if is_refs_title(t):
            in_body=False
            continue
        if lvl==1: in_body=True
        if not in_body or not t or lvl or is_caption(t) or is_ref(t) or has_drawing(p): continue
        if p.getparent() is not body: continue
        fr=first_run(p); ia=ind(p)
        if fr and not ok_body_font(fr): fail(fails,'body',f'body font wrong idx={idx}: {t[:60]} props={fr}')
        if jc(p)!='both': fail(fails,'body',f'body align not justified idx={idx}: {t[:60]} jc={jc(p)}')
        if ia.get('firstLineChars')!='200' and ia.get('firstLine') not in ('480',): warn(warns,'body',f'body indent suspicious idx={idx}: {t[:60]} ind={ia}')

    table_count=0
    for tbl in root.findall('.//w:tbl',NS):
        table_count+=1
        tb=tbl.find('./w:tblPr/w:tblBorders',NS)
        if tb is None: fail(fails,'tables',f'table {table_count} missing borders'); continue
        def bv(edge):
            e=tb.find(q(edge)); return None if e is None else e.get(q('val'))
        if bv('top')!='single' or bv('bottom')!='single': fail(fails,'tables',f'table {table_count} missing top/bottom')
        for edge in ('left','right','insideH','insideV'):
            if bv(edge) not in ('nil','none'): fail(fails,'tables',f'table {table_count} non-open border {edge}={bv(edge)}')
        for ri,row in enumerate(tbl.findall('./w:tr',NS),1):
            trPr=row.find(q('trPr'))
            if trPr is None or trPr.find(q('cantSplit')) is None: warn(warns,'tables',f'table {table_count} row {ri} missing cantSplit')
            for cell in row.findall('./w:tc',NS):
                for p in cell.findall('.//w:p',NS):
                    if text(p) and jc(p)!='center': fail(fails,'tables',f'table {table_count} cell not centered: {text(p)[:40]}')
                    sp=spacing(p)
                    if text(p) and sp.get('lineRule')!='auto': fail(fails,'tables',f'table {table_count} cell not single/auto spacing: {text(p)[:40]} spacing={sp}')
                    for r in p.findall('./w:r',NS):
                        rp=run_props(r)
                        if text(p) and not ok_table_font(rp):
                            fail(fails,'tables',f'table {table_count} font wrong: {text(p)[:40]} props={rp}'); break
    metrics['tables']=table_count

    in_body=False; pic_count=0
    for idx,p in enumerate(paras):
        t=text(p); lvl=h_level(t)
        if lvl==1: in_body=True
        if not in_body or not has_drawing(p): continue
        pic_count+=1
        if has_anchor(p): fail(fails,'figures',f'body picture has floating anchor idx={idx}')
        if t: fail(fails,'figures',f'body picture paragraph contains text idx={idx}: {t[:60]}')
        if not has_inline(p): warn(warns,'figures',f'body picture has no inline idx={idx}')
    metrics['body_pictures']=pic_count

    if markers['refs']:
        in_refs=False; ref_count=0
        for p in paras:
            t=text(p)
            if is_refs_title(t): in_refs=True; continue
            if in_refs and t:
                if h_level(t)==1: break
                if is_ref(t): ref_count+=1
        metrics['references']=ref_count
        if ref_count==0: warn(warns,'references','reference title exists but no [n] entries detected')

    print('FILE:',path)
    print('METRICS:',json.dumps(metrics,ensure_ascii=False))
    print('WARNING_COUNT:',len(warns))
    for w in warns[:200]: print('WARN',w['module'],w['msg'])
    print('FAILURE_COUNT:',len(fails))
    for f in fails[:240]: print('FAIL',f['module'],f['msg'])
    print('RESULT:','通过' if not fails else '未通过')
    return 0 if not fails else 1

if __name__=='__main__': raise SystemExit(audit(sys.argv[1]))
