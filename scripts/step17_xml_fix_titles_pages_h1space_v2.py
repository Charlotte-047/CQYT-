from zipfile import ZipFile, ZIP_DEFLATED
from pathlib import Path
from lxml import etree as ET
import sys, tempfile, shutil, re

W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
REL='{http://schemas.openxmlformats.org/package/2006/relationships}'
CT='{http://schemas.openxmlformats.org/package/2006/content-types}'
NS={'w':W.strip('{}'),'r':R.strip('{}'),'rel':REL.strip('{}'),'ct':CT.strip('{}')}
H1_TEXTS={'致谢','参考文献','摘  要','ABSTRACT','目  录'}
HEADER_TEXT='重庆移通学院20××届毕业论文（设计）'

def q(t): return f'{W}{t}'
def rq(t): return f'{R}{t}'
def text(p): return ''.join(t.text or '' for t in p.findall('.//w:t',NS)).strip()
def style(p):
    ppr=p.find(q('pPr')); ps=None if ppr is None else ppr.find(q('pStyle'))
    return '' if ps is None else ps.get(q('val')) or ''
def ensure(parent,tag,first=False):
    n=parent.find(q(tag))
    if n is None:
        n=ET.Element(q(tag))
        if first: parent.insert(0,n)
        else: parent.append(n)
    return n

def set_run(r,east='宋体',asc='Times New Roman',size=24,bold=False,color='000000'):
    rpr=ensure(r,'rPr',True)
    for c in list(rpr):
        if c.tag in (q('rFonts'),q('sz'),q('szCs'),q('b'),q('bCs'),q('color'),q('u')): rpr.remove(c)
    rf=ET.SubElement(rpr,q('rFonts')); rf.set(q('eastAsia'),east); rf.set(q('ascii'),asc); rf.set(q('hAnsi'),asc); rf.set(q('cs'),asc)
    b=ET.SubElement(rpr,q('b')); b.set(q('val'),'1' if bold else '0')
    bcs=ET.SubElement(rpr,q('bCs')); bcs.set(q('val'),'1' if bold else '0')
    sz=ET.SubElement(rpr,q('sz')); sz.set(q('val'),str(size))
    szcs=ET.SubElement(rpr,q('szCs')); szcs.set(q('val'),str(size))
    c=ET.SubElement(rpr,q('color')); c.set(q('val'),color)

def set_all_runs(p,east='宋体',asc='Times New Roman',size=24,bold=False,color='000000'):
    rs=p.findall('./w:r',NS)
    if not rs:
        r=ET.SubElement(p,q('r')); ET.SubElement(r,q('t')).text=''; rs=[r]
    for r in rs: set_run(r,east,asc,size,bold,color)

def set_ppr(p,align='left',first='0',line='440'):
    ppr=ensure(p,'pPr',True)
    jc=ensure(ppr,'jc'); jc.set(q('val'),align)
    sp=ensure(ppr,'spacing'); sp.set(q('before'),'0'); sp.set(q('after'),'0'); sp.set(q('line'),line); sp.set(q('lineRule'),'exact')
    ind=ensure(ppr,'ind')
    for a in ('firstLine','left','hanging','right','firstLineChars','leftChars','hangingChars','rightChars'):
        ind.attrib.pop(q(a),None)
    if first is not None: ind.set(q('firstLine'),first)

def remove_pagebreak_before(p):
    ppr=p.find(q('pPr'))
    if ppr is not None:
        for pb in ppr.findall(q('pageBreakBefore')): ppr.remove(pb)

def remove_all_br(p):
    for br in list(p.findall('.//w:br',NS)):
        parent=br.getparent(); parent.remove(br)

def has_page_break(p): return bool(p.findall('.//w:br',NS))
def add_page_break(p):
    if has_page_break(p): return
    r=ET.SubElement(p,q('r')); set_run(r,'宋体','Times New Roman',24,False,'000000')
    br=ET.SubElement(r,q('br')); br.set(q('type'),'page')

def is_blank(p): return p is not None and p.tag==q('p') and text(p)==''
def make_blank():
    p=ET.Element(q('p')); set_ppr(p,'left','0','440'); set_all_runs(p,'宋体','Times New Roman',24,False,'000000'); return p

def format_blank(p):
    ppr=ensure(p,'pPr',True)
    ps=ppr.find(q('pStyle'))
    if ps is not None: ppr.remove(ps)
    set_ppr(p,'left','0','440')
    # A truly empty paragraph can be collapsed/eaten by WPS at a page top. Use a visible-size
    # whitespace run so the required blank line renders, while strip()/text() still treats it blank.
    for r in list(p.findall('./w:r',NS)):
        p.remove(r)
    r=ET.SubElement(p,q('r')); set_run(r,'宋体','Times New Roman',24,False,'000000')
    t=ET.SubElement(r,q('t')); t.text=' '; t.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
    remove_pagebreak_before(p); remove_all_br(p)

def looks_like_body_sentence(t):
    # Headings may contain ：、（）/ etc. Only reject obvious long sentences.
    return len(t) > 70 or bool(re.search(r'[。；;]$', t))

def heading_level(p):
    t = text(p)
    if looks_like_body_sentence(t):
        return None
    # Numbering is authoritative when present. Match deeper heading first so
    # 1.1.1 -> H3, 1.1 -> H2, 1 -> H1. Do not trust stale Word styles.
    if re.match(r'^\d+\.\d+\.\d+\s*\S.{0,70}$', t): return 3
    if re.match(r'^\d+\.\d+\s*\S.{0,65}$', t): return 2
    if t in H1_TEXTS or re.match(r'^\d+\s+\S.{0,55}$', t): return 1
    return None

def is_toc_heading_like_text(t):
    return bool(re.match(r'^(\d+(?:\.\d+)*\s+.+|致谢|参考文献)\s*\d+$', t)) or ('\t' in t)

def format_h1(p):
    ppr=ensure(p,'pPr',True); ps=ensure(ppr,'pStyle',True); ps.set(q('val'),'3')
    set_ppr(p,'center','0','440')
    # Paragraph before/after is still inserted. SpaceBefore/After is an extra WPS-safe fallback.
    sp=ensure(ppr,'spacing'); sp.set(q('before'),'440'); sp.set(q('after'),'440'); sp.set(q('line'),'440'); sp.set(q('lineRule'),'exact')
    # Real new-page handling should live on the H1 itself for Word/WPS stability.
    pbb=ensure(ppr,'pageBreakBefore')
    pbb.set(q('val'),'1')
    set_all_runs(p,'黑体','Times New Roman',32,False,'000000'); remove_all_br(p)

def format_h2_h3(p):
    lvl = heading_level(p)
    ppr=ensure(p,'pPr',True); ps=ensure(ppr,'pStyle',True)
    if lvl == 2:
        ps.set(q('val'),'4')
        set_ppr(p,'left','0','440'); set_all_runs(p,'黑体','Times New Roman',30,False,'000000'); remove_pagebreak_before(p); remove_all_br(p)
    elif lvl == 3:
        ps.set(q('val'),'5')
        set_ppr(p,'left','0','440'); set_all_runs(p,'黑体','Times New Roman',28,False,'000000'); remove_pagebreak_before(p); remove_all_br(p)

def patch_heading_run_colors(tmp):
    docxml=tmp/'word'/'document.xml'; tree=ET.parse(str(docxml)); root=tree.getroot(); changed=0
    body=root.find(q('body'))
    in_toc=False
    for p in list(body):
        if p.tag!=q('p'):
            continue
        t=text(p)
        if t in ('目  录','目    录'):
            in_toc=True
            continue
        if in_toc:
            if re.match(r'^1\s+\S+', t) and not re.search(r'\d+$', t):
                in_toc=False
            else:
                continue
        lvl = heading_level(p)
        if lvl in (1,2,3):
            if lvl == 1: format_h1(p)
            else: format_h2_h3(p)
            changed += 1
    tree.write(str(docxml),xml_declaration=True,encoding='utf-8',standalone='yes')
    return changed

def copy_sectPr(src): return ET.fromstring(ET.tostring(src))

def patch_styles(tmp):
    styles=tmp/'word'/'styles.xml'
    if not styles.exists(): return 0
    tree=ET.parse(str(styles)); root=tree.getroot(); changed=0
    specs={'3':('32','center'),'4':('30','left'),'5':('28','left'),'6':('24','left'),'Heading1':('32','center'),'Heading2':('30','left'),'Heading3':('28','left'),'Heading4':('24','left')}
    for sid,(size,align) in specs.items():
        st=root.find(f".//w:style[@w:styleId='{sid}']",NS)
        if st is None: continue
        ppr=ensure(st,'pPr'); jc=ensure(ppr,'jc'); jc.set(q('val'),align)
        sp=ensure(ppr,'spacing'); sp.set(q('before'),'0'); sp.set(q('after'),'0'); sp.set(q('line'),'440'); sp.set(q('lineRule'),'exact')
        ind=ensure(ppr,'ind')
        for a in list(ind.attrib): ind.attrib.pop(a,None)
        ind.set(q('firstLine'),'0')
        rpr=ensure(st,'rPr')
        for c in list(rpr):
            if c.tag in (q('rFonts'),q('sz'),q('szCs'),q('b'),q('bCs'),q('color'),q('u')): rpr.remove(c)
        rf=ET.SubElement(rpr,q('rFonts')); rf.set(q('eastAsia'),'黑体'); rf.set(q('ascii'),'Times New Roman'); rf.set(q('hAnsi'),'Times New Roman'); rf.set(q('cs'),'Times New Roman')
        b=ET.SubElement(rpr,q('b')); b.set(q('val'),'0')
        bcs=ET.SubElement(rpr,q('bCs')); bcs.set(q('val'),'0')
        sz=ET.SubElement(rpr,q('sz')); sz.set(q('val'),size)
        szcs=ET.SubElement(rpr,q('szCs')); szcs.set(q('val'),size)
        c=ET.SubElement(rpr,q('color')); c.set(q('val'),'000000')
        changed+=1
    tree.write(str(styles),xml_declaration=True,encoding='utf-8',standalone='yes')
    return changed

def ensure_body_section_before_first_h1(body):
    # Preserve original front matter/TOC/abstract sectioning. Do not infer by a
    # hard-coded '1 绪论' marker, because source chapter names and numbering vary.
    return False

def patch_h1_spacing(tmp):
    docxml=tmp/'word'/'document.xml'; tree=ET.parse(str(docxml)); root=tree.getroot(); body=root.find(q('body'))
    ensure_body_section_before_first_h1(body)
    children=list(body); i=0; touched=0; in_toc=False
    while i < len(children):
        node=children[i]
        if node.tag==q('p'):
            t=text(node)
            if t in ('目  录','目    录'):
                # TOC title must never enter H1 page-break formatting pass.
                in_toc=True
                i+=1
                continue
            if in_toc:
                # Hard TOC fence: after TOC title, do not touch any heading-like entries until
                # the first pure body H1 without trailing page number appears.
                if re.match(r'^1\s+\S+', t) and not re.search(r'\d+$', t):
                    in_toc=False
                else:
                    i+=1
                    continue
            if heading_level(node)==1 and not in_toc:
                format_h1(node)
                idx=list(body).index(node)
                if idx==0 or not is_blank(list(body)[idx-1]):
                    body.insert(idx,make_blank()); idx+=1; node=list(body)[idx]
                format_blank(list(body)[idx-1])
                # Keep explicit page break off the surrounding blanks; H1 itself carries pageBreakBefore.
                if idx>=2:
                    before_blank=list(body)[idx-2]
                    if before_blank.tag==q('p'):
                        remove_pagebreak_before(before_blank)
                        remove_all_br(before_blank)
                if idx+1>=len(list(body)) or not is_blank(list(body)[idx+1]):
                    body.insert(idx+1,make_blank())
                format_blank(list(body)[idx+1])
                touched+=1
                children=list(body); i=idx+2; continue
        i+=1
    tree.write(str(docxml),xml_declaration=True,encoding='utf-8',standalone='yes')
    return touched

def next_rid(root):
    nums=[]
    for rel in root:
        rid=rel.get('Id','')
        if rid.startswith('rId') and rid[3:].isdigit(): nums.append(int(rid[3:]))
    return max(nums+[0])+1

def ensure_override(tmp,part,ctype):
    ctp=tmp/'[Content_Types].xml'; tree=ET.parse(str(ctp)); root=tree.getroot(); tag=f'{CT}Override'
    for o in root.findall(tag):
        if o.get('PartName')==part:
            o.set('ContentType',ctype); tree.write(str(ctp),xml_declaration=True,encoding='utf-8',standalone='yes'); return
    ET.SubElement(root,tag,PartName=part,ContentType=ctype); tree.write(str(ctp),xml_declaration=True,encoding='utf-8',standalone='yes')

def header_xml():
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="{W.strip('{}')}" xmlns:r="{R.strip('{}')}"><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:rFonts w:eastAsia="宋体" w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:color w:val="000000"/></w:rPr><w:t>{HEADER_TEXT}</w:t></w:r></w:p></w:hdr>'''
def footer_blank_xml(): return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="{W.strip('{}')}" xmlns:r="{R.strip('{}')}"><w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p></w:ftr>'''
def footer_page_xml(fmt=None):
    instr=' PAGE'
    if fmt=='roman': instr=' PAGE \\* ROMAN'
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="{W.strip('{}')}" xmlns:r="{R.strip('{}')}"><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:rFonts w:eastAsia="宋体" w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:color w:val="000000"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">{instr}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>'''

def clear_refs(sect):
    for old in sect.findall(q('headerReference'))+sect.findall(q('footerReference')): sect.remove(old)
def add_ref(sect,tag,rid):
    el=ET.Element(q(tag)); el.set(q('type'),'default'); el.set(rq('id'),rid); sect.insert(0 if tag=='headerReference' else 1,el)
def set_pg_start(sect,fmt=None,start=None):
    pg=sect.find(q('pgNumType'))
    if pg is None: pg=ET.SubElement(sect,q('pgNumType'))
    if fmt: pg.set(q('fmt'),fmt)
    else: pg.attrib.pop(q('fmt'),None)
    if start is not None: pg.set(q('start'),str(start))
    else: pg.attrib.pop(q('start'),None)

def ensure_final_body_sectPr(body):
    last = list(body)[-1] if len(body) else None
    if last is not None and last.tag == q('sectPr'):
        return False
    template = None
    sects = body.findall('.//w:sectPr', NS)
    if sects:
        template = copy_sectPr(sects[-1])
    else:
        template = ET.Element(q('sectPr'))
    body.append(template)
    return True


def patch_headers_footers(tmp):
    docxml_pre=tmp/'word'/'document.xml'
    pre_tree=ET.parse(str(docxml_pre)); pre_root=pre_tree.getroot(); pre_body=pre_root.find(q('body'))
    ensure_final_body_sectPr(pre_body)
    pre_tree.write(str(docxml_pre),xml_declaration=True,encoding='utf-8',standalone='yes')

    word=tmp/'word'; relp=word/'_rels'/'document.xml.rels'; rtree=ET.parse(str(relp)); rroot=rtree.getroot(); n=next_rid(rroot)
    ids={'hdr':f'rId{n}','blank':f'rId{n+1}','roman':f'rId{n+2}','arabic':f'rId{n+3}'}
    (word/'header_step17v2.xml').write_text(header_xml(),encoding='utf-8')
    (word/'footer_blank_step17v2.xml').write_text(footer_blank_xml(),encoding='utf-8')
    (word/'footer_roman_step17v2.xml').write_text(footer_page_xml('roman'),encoding='utf-8')
    (word/'footer_arabic_step17v2.xml').write_text(footer_page_xml(None),encoding='utf-8')
    for part,ctype in [('/word/header_step17v2.xml','application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'),('/word/footer_blank_step17v2.xml','application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'),('/word/footer_roman_step17v2.xml','application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'),('/word/footer_arabic_step17v2.xml','application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml')]: ensure_override(tmp,part,ctype)
    ET.SubElement(rroot,f'{REL}Relationship',Id=ids['hdr'],Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',Target='header_step17v2.xml')
    ET.SubElement(rroot,f'{REL}Relationship',Id=ids['blank'],Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',Target='footer_blank_step17v2.xml')
    ET.SubElement(rroot,f'{REL}Relationship',Id=ids['roman'],Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',Target='footer_roman_step17v2.xml')
    ET.SubElement(rroot,f'{REL}Relationship',Id=ids['arabic'],Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',Target='footer_arabic_step17v2.xml')
    rtree.write(str(relp),xml_declaration=True,encoding='utf-8',standalone='yes')
    docxml=word/'document.xml'; tree=ET.parse(str(docxml)); root=tree.getroot(); sects=root.findall(f'.//{q("sectPr")}')
    # After split, expected: 1 cover, 2 declaration, 3 CN abstract, 4 EN abstract, 5 TOC, 6 body.
    # But tolerate variants by using last section as body and section containing TOC as no-page.
    for idx,sect in enumerate(sects,1):
        clear_refs(sect)
        if idx==1 or idx==2 or idx==5:
            add_ref(sect,'footerReference',ids['blank']); set_pg_start(sect,None,None)
        elif idx in (3,4):
            add_ref(sect,'headerReference',ids['hdr']); add_ref(sect,'footerReference',ids['roman']); set_pg_start(sect,'roman',1 if idx==3 else None)
        else:
            add_ref(sect,'headerReference',ids['hdr']); add_ref(sect,'footerReference',ids['arabic']); set_pg_start(sect,None,1 if idx==6 else None)
    tree.write(str(docxml),xml_declaration=True,encoding='utf-8',standalone='yes')
    return len(sects)

def process(src,out):
    src=Path(src); out=Path(out); out.parent.mkdir(parents=True,exist_ok=True); tmp=Path(tempfile.mkdtemp(prefix='step17v2_'))
    try:
        with ZipFile(src) as z: z.extractall(tmp)
        styles=patch_styles(tmp); h1=patch_h1_spacing(tmp); heading_runs=patch_heading_run_colors(tmp); sections=patch_headers_footers(tmp)
        repack=out.with_suffix('.tmp.docx')
        with ZipFile(repack,'w',ZIP_DEFLATED) as z:
            for f in tmp.rglob('*'):
                if f.is_file(): z.write(f,f.relative_to(tmp))
        shutil.move(str(repack),str(out))
        print('OUTPUT:',out); print('STYLES_FIXED:',styles); print('H1_SPACING_FIXED:',h1); print('HEADING_RUNS_FIXED:',heading_runs); print('SECTIONS:',sections)
    finally: shutil.rmtree(tmp,ignore_errors=True)
if __name__=='__main__':
    if len(sys.argv)!=3: raise SystemExit('usage: python step17_xml_fix_titles_pages_h1space_v2.py in.docx out.docx')
    process(sys.argv[1],sys.argv[2])
