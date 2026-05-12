from zipfile import ZipFile
from pathlib import Path
from lxml import etree as ET
import sys,re,json

W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
WP='{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
NS={'w':W.strip('{}'),'wp':WP.strip('{}')}
def q(t): return f'{W}{t}'
def text(n): return ''.join(t.text or '' for t in n.findall('.//w:t',NS)).strip()
def style(p):
    ppr=p.find(q('pPr')); ps=None if ppr is None else ppr.find(q('pStyle'))
    return '' if ps is None else (ps.get(q('val')) or '')
def ppr(p): return p.find(q('pPr'))
def jc(p):
    pp=ppr(p); n=None if pp is None else pp.find(q('jc'))
    return None if n is None else n.get(q('val'))
def spacing(p):
    pp=ppr(p); n=None if pp is None else pp.find(q('spacing'))
    return {} if n is None else {k.split('}')[-1]:v for k,v in n.attrib.items()}
def ind(p):
    pp=ppr(p); n=None if pp is None else pp.find(q('ind'))
    return {} if n is None else {k.split('}')[-1]:v for k,v in n.attrib.items()}
def run_props(r):
    rpr=r.find(q('rPr'))
    rf=None if rpr is None else rpr.find(q('rFonts'))
    sz=None if rpr is None else rpr.find(q('sz'))
    b=None if rpr is None else rpr.find(q('b'))
    color=None if rpr is None else rpr.find(q('color'))
    return {
        'east':None if rf is None else rf.get(q('eastAsia')),
        'ascii':None if rf is None else rf.get(q('ascii')),
        'hAnsi':None if rf is None else rf.get(q('hAnsi')),
        'size':None if sz is None else sz.get(q('val')),
        'bold':False if b is None else b.get(q('val')) not in ('0','false','False'),
        'color':None if color is None else color.get(q('val')),
    }
def first_run_props(p):
    r=p.find('./w:r',NS)
    return {} if r is None else run_props(r)
def has_pagebreak_before(p):
    pp=ppr(p); n=None if pp is None else pp.find(q('pageBreakBefore'))
    return n is not None
def has_br_page(p): return bool(p.findall('.//w:br[@w:type="page"]',NS))
def has_drawing(p): return bool(p.findall('.//w:drawing',NS))
def has_anchor(p): return bool(p.findall('.//wp:anchor',NS))
def has_inline(p): return bool(p.findall('.//wp:inline',NS))
def is_blank(n): return n.tag==q('p') and text(n)=='' and not has_drawing(n)
def looks_like_body_sentence(t): return len(t)>70 or bool(re.search(r'[。；;]$',t))
def safe_h_level(t):
    if looks_like_body_sentence(t): return None
    if t in ('致谢','参考文献'): return 1
    if re.match(r'^\d+\.\d+\.\d+\s*\S.{0,70}$',t): return 3
    if re.match(r'^\d+\.\d+\s*\S.{0,65}$',t): return 2
    if re.match(r'^\d+\s+\S.{0,55}$',t): return 1
    return None

def fail(fails,msg): fails.append(msg)
def border_is_none(b): return (not b) or b.get('val') in (None,'nil','none') or b.get('sz')=='0'
def battrs(n): return {} if n is None else {k.split('}')[-1]:v for k,v in n.attrib.items()}
def cell_border_map(tc):
    tcPr=tc.find(q('tcPr')); b=None if tcPr is None else tcPr.find(q('tcBorders'))
    out={}
    if b is None: return out
    for name in ('top','left','bottom','right','insideH','insideV','tl2br','tr2bl'):
        n=b.find(q(name))
        if n is not None: out[name]=battrs(n)
    return out

def verify(path):
    path=Path(path)
    fails=[]; warns=[]; metrics={}
    with ZipFile(path) as z:
        root=ET.fromstring(z.read('word/document.xml'))
    body=root.find(q('body')); nodes=list(body); paras=[n for n in nodes if n.tag==q('p')]

    # 1. TOC must not contain page breaks or heading styles.
    in_toc=False; toc_lines=0
    for p in paras:
        t=text(p)
        if t in ('目  录','目    录'):
            in_toc=True; continue
        if in_toc:
            if safe_h_level(t)==1 and not re.search(r'\d$',t):
                in_toc=False
                continue
            if t:
                toc_lines += 1
                if style(p) in ('3','4','5','Heading1','Heading2','Heading3'):
                    fail(fails,f'TOC line has heading style: {t[:50]} style={style(p)}')
                # TOC is protected by default; preserve source page/section marks inside it.
    metrics['toc_lines_checked']=toc_lines

    # 2. Safe headings only; long body-like text must not be heading style or heading-sized bold font.
    heading_count=0; restored_risk=0; in_toc=False
    for p in paras:
        t=text(p)
        if not t: continue
        if re.sub(r'\s+','',t or '')=='\u76ee\u5f55':
            in_toc=True
            continue
        if in_toc:
            if re.search(r'\d\s*$', t):
                continue
            if safe_h_level(t)==1:
                in_toc=False
            else:
                continue
        lvl=safe_h_level(t); st=style(p); rp=first_run_props(p)
        if lvl:
            heading_count += 1
            exp_size={1:'32',2:'30',3:'28'}[lvl]
            if st not in (str(1+lvl), f'Heading{lvl}'):
                fail(fails,f'safe heading wrong style: {t[:50]} style={st}')
            if rp.get('east')!='\u9ed1\u4f53' or rp.get('size')!=exp_size:
                fail(fails,f'safe heading wrong font: {t[:50]} props={rp}')
        elif looks_like_body_sentence(t):
            if st in ('3','4','5','Heading1','Heading2','Heading3') or (rp.get('east')=='\u9ed1\u4f53' and rp.get('size') in ('32','30','28')):
                fail(fails,f'body-like paragraph polluted as heading: {t[:80]} style={st} props={rp}')
                restored_risk += 1
    metrics['safe_headings_checked']=heading_count
    metrics['body_heading_pollution_hits']=restored_risk

    # 3. Tables: t0 three-line table + centered table body + 五号 + single spacing.
    table_count=0
    for tbl in root.findall('.//w:tbl',NS):
        table_count += 1
        rows=tbl.findall('./w:tr',NS)
        if not rows:
            fail(fails,f'table {table_count} has no rows')
            continue
        for ri,row in enumerate(rows):
            trPr=row.find(q('trPr'))
            if trPr is None or trPr.find(q('cantSplit')) is None: fail(fails,f'table {table_count} row {ri+1} missing cantSplit')
            for ci,cell in enumerate(row.findall('./w:tc',NS),1):
                cb=cell_border_map(cell)
                for edge in ('left','right','insideH','insideV','tl2br','tr2bl'):
                    if not border_is_none(cb.get(edge,{})): fail(fails,f'table {table_count} cell {ri+1},{ci} unexpected {edge} border')
                top=cb.get('top',{}); bottom=cb.get('bottom',{})
                if ri==0:
                    if top.get('val')!='single' or top.get('sz')!='12': fail(fails,f'table {table_count} header top not 1.5pt')
                    if bottom.get('val')!='single' or bottom.get('sz')!='6': fail(fails,f'table {table_count} header bottom not 0.75pt')
                elif ri==len(rows)-1:
                    if not border_is_none(top): fail(fails,f'table {table_count} last row unexpected top border')
                    if bottom.get('val')!='single' or bottom.get('sz')!='12': fail(fails,f'table {table_count} last row bottom not 1.5pt')
                else:
                    if not border_is_none(top) or not border_is_none(bottom): fail(fails,f'table {table_count} body row has extra horizontal border')
                for p in cell.findall('.//w:p',NS):
                    if jc(p)!='center': fail(fails,f'table {table_count} cell paragraph not centered: {text(p)[:30]} jc={jc(p)}')
                    sp=spacing(p)
                    if sp.get('lineRule')!='auto': fail(fails,f'table {table_count} cell not single/auto spacing: {text(p)[:30]} spacing={sp}')
                    for r in p.findall('./w:r',NS):
                        rp=run_props(r)
                        if rp.get('east')!='宋体' or rp.get('ascii')!='Times New Roman' or rp.get('size')!='21':
                            fail(fails,f'table {table_count} run font wrong: {text(p)[:30]} props={rp}')
                            break
    metrics['tables_checked']=table_count

    # 4. Pictures: only check body pictures after the first real H1; front cover/frontmatter
    # often uses anchored shapes/text boxes and should not fail body media validation.
    pic_count=0; in_body=False
    for p in paras:
        t=text(p)
        if safe_h_level(t)==1 and style(p) in ('3','Heading1'):
            in_body=True
        if not in_body:
            continue
        if has_drawing(p):
            pic_count += 1
            if has_anchor(p): fail(fails,f'body picture paragraph still has anchor near: {text(p)[:30]}')
            if text(p): fail(fails,f'body picture paragraph contains text: {text(p)[:50]}')
            if not has_inline(p): warns.append(f'body picture paragraph has drawing but no inline: {text(p)[:30]}')
    metrics['body_pictures_checked']=pic_count

    # 5. H1 page break should apply to real H1 only, not TOC.
    in_toc=False
    for p in paras:
        t=text(p)
        if re.sub(r'\s+','',t or '')=='\u76ee\u5f55':
            in_toc=True; continue
        if in_toc:
            if re.search(r'\d\s*$', t):
                continue
            if safe_h_level(t)==1:
                in_toc=False
            else:
                continue
        lvl=safe_h_level(t)
        if lvl==1 and style(p) in ('3','Heading1'):
            if not has_pagebreak_before(p): fail(fails,f'H1 missing pageBreakBefore: {t}')

    print('FILE:', path)
    print('METRICS:', json.dumps(metrics, ensure_ascii=False))
    print('WARNING_COUNT:', len(warns))
    for w in warns[:30]: print('WARN:', w)
    print('FAILURE_COUNT:', len(fails))
    for f in fails[:120]: print('FAIL:', f)
    print('RESULT:', '通过' if not fails else '未通过')
    return 0 if not fails else 1

if __name__=='__main__':
    raise SystemExit(verify(sys.argv[1]))
