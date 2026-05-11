from zipfile import ZipFile, ZIP_DEFLATED
from pathlib import Path
from lxml import etree as ET
import tempfile, shutil, sys, re

W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
NS={'w':W.strip('{}'),'r':R.strip('{}')}
FONT_HEI=chr(0x9ed1)+chr(0x4f53)
FONT_SONG=chr(0x5b8b)+chr(0x4f53)
TXT_ABS=chr(0x6458)+chr(0x8981)
TXT_CN_ABS=chr(0x4e2d)+chr(0x6587)+TXT_ABS
TXT_TOC=chr(0x76ee)+chr(0x5f55)

def q(t): return f'{W}{t}'
def ensure(parent,tag,first=False):
    n=parent.find(q(tag))
    if n is None:
        n=ET.Element(q(tag))
        if first: parent.insert(0,n)
        else: parent.append(n)
    return n

def clear_children(parent, tags):
    for tag in tags:
        for n in list(parent.findall(q(tag))):
            parent.remove(n)

def set_border(parent,edge,val='single',sz='8'):
    e=parent.find(q(edge))
    if e is None: e=ET.SubElement(parent,q(edge))
    e.attrib.clear()
    e.set(q('val'),val); e.set(q('sz'),sz); e.set(q('space'),'0'); e.set(q('color'),'000000')

def nil_border(parent,edge):
    e=parent.find(q(edge))
    if e is None: e=ET.SubElement(parent,q(edge))
    e.attrib.clear(); e.set(q('val'),'nil')

def none_border(parent,edge):
    e=parent.find(q(edge))
    if e is None: e=ET.SubElement(parent,q(edge))
    e.attrib.clear(); e.set(q('val'),'none'); e.set(q('color'),'auto'); e.set(q('sz'),'0'); e.set(q('space'),'0')

def fix_sections(root):
    sects=root.findall('.//w:sectPr',NS)
    for i,sect in enumerate(sects,1):
        pg=sect.find(q('pgNumType'))
        if i in (3,4):
            if pg is None:
                pg=ET.SubElement(sect,q('pgNumType'))
            pg.attrib.clear()
            pg.set(q('fmt'),'roman')
            if i==3:
                pg.set(q('start'),'1')
        elif i>=5:
            if pg is None:
                pg=ET.SubElement(sect,q('pgNumType'))
            pg.attrib.clear()
            pg.set(q('fmt'),'decimal')
            if i==6:
                pg.set(q('start'),'1')

def normalize_table_paragraphs(cell):
    for p in cell.findall('.//w:p',NS):
        ppr=ensure(p,'pPr',True)
        ps=ppr.find(q('pStyle'))
        if ps is not None:
            ppr.remove(ps)
        jc=ppr.find(q('jc'))
        if jc is None:
            jc=ET.SubElement(ppr,q('jc'))
        # User requested table content centered.
        jc.set(q('val'),'center')
        ind=ensure(ppr,'ind')
        for a in ('firstLine','left','hanging','right','firstLineChars','leftChars','hangingChars','rightChars'):
            ind.attrib.pop(q(a),None)
        sp=ensure(ppr,'spacing')
        sp.attrib.clear()
        sp.set(q('before'),'0')
        sp.set(q('after'),'0')
        sp.set(q('lineRule'),'auto')
        # normalize runs to 宋体/Times New Roman 五号
        for r in p.findall('./w:r',NS):
            rpr=ensure(r,'rPr',True)
            for c in list(rpr):
                if c.tag in (q('rFonts'),q('sz'),q('szCs'),q('b'),q('bCs'),q('color'),q('u')):
                    rpr.remove(c)
            rf=ET.SubElement(rpr,q('rFonts'))
            rf.set(q('eastAsia'),'宋体'); rf.set(q('ascii'),'Times New Roman'); rf.set(q('hAnsi'),'Times New Roman'); rf.set(q('cs'),'Times New Roman')
            b=ET.SubElement(rpr,q('b')); b.set(q('val'),'0')
            bcs=ET.SubElement(rpr,q('bCs')); bcs.set(q('val'),'0')
            sz=ET.SubElement(rpr,q('sz')); sz.set(q('val'),'21')
            szcs=ET.SubElement(rpr,q('szCs')); szcs.set(q('val'),'21')
            color=ET.SubElement(rpr,q('color')); color.set(q('val'),'000000')

def fix_tables(root):
    count=0
    for tbl in root.findall('.//w:tbl',NS):
        tblPr=ensure(tbl,'tblPr',True)
        clear_children(tblPr,['tblStyle','tblLook'])
        tb=tblPr.find(q('tblBorders'))
        if tb is None: tb=ET.SubElement(tblPr,q('tblBorders'))
        else:
            for c in list(tb): tb.remove(c)
        set_border(tb,'top','single','12')
        set_border(tb,'bottom','single','12')
        nil_border(tb,'left'); nil_border(tb,'right'); nil_border(tb,'insideH'); nil_border(tb,'insideV')
        rows=tbl.findall('./w:tr',NS)
        for ri,row in enumerate(rows):
            trPr=ensure(row,'trPr',True)
            if trPr.find(q('cantSplit')) is None:
                ET.SubElement(trPr,q('cantSplit'))
            for cell in row.findall('./w:tc',NS):
                tcPr=ensure(cell,'tcPr',True)
                cb=tcPr.find(q('tcBorders'))
                if cb is None: cb=ET.SubElement(tcPr,q('tcBorders'))
                else:
                    for c in list(cb): cb.remove(c)
                nil_border(cb,'left'); nil_border(cb,'right'); nil_border(cb,'insideH'); nil_border(cb,'insideV')
                if ri==0:
                    set_border(cb,'top','single','12')
                    set_border(cb,'bottom','single','12')
                else:
                    nil_border(cb,'top')
                    if ri==len(rows)-1:
                        set_border(cb,'bottom','single','12')
                    else:
                        nil_border(cb,'bottom')
                normalize_table_paragraphs(cell)
        count += 1
    return count

def text(p): return ''.join(t.text or '' for t in p.findall('.//w:t',NS)).strip()
def style(p):
    ppr=p.find(q('pPr')); ps=None if ppr is None else ppr.find(q('pStyle'))
    return '' if ps is None else (ps.get(q('val')) or '')
def looks_like_body_sentence(t): return len(t)>70 or bool(re.search(r'[。；;]$',t))
def safe_heading_level(t):
    if looks_like_body_sentence(t): return None
    if t in ('致谢','参考文献'): return 1
    if re.match(r'^\d+\.\d+\.\d+\s*\S.{0,70}$',t): return 3
    if re.match(r'^\d+\.\d+\s*\S.{0,65}$',t): return 2
    if re.match(r'^\d+\s+\S.{0,55}$',t): return 1
    return None

def norm_space(t):
    return re.sub(r'\s+','',t or '')

def is_cn_abs_title(t):
    n=norm_space(t)
    return n in (TXT_ABS,TXT_CN_ABS) or (chr(0x6458) in n and chr(0x8981) in n and len(n)<=6)

def is_en_abs_title(t):
    return norm_space(t).upper() == 'ABSTRACT' or (norm_space(t).startswith('Abstract') and len(norm_space(t))<=12)

def is_toc_title(t):
    return norm_space(t) == TXT_TOC

def make_para(text_value='', style_val=None):
    p=ET.Element(q('p'))
    ppr=ET.SubElement(p,q('pPr'))
    if style_val:
        ps=ET.SubElement(ppr,q('pStyle')); ps.set(q('val'),style_val)
    r=ET.SubElement(p,q('r'))
    t=ET.SubElement(r,q('t'))
    t.text=text_value
    if text_value.startswith(' ') or text_value.endswith(' '):
        t.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
    return p

def replace_para_text(p,new_text):
    for r in list(p.findall('./w:r',NS)):
        p.remove(r)
    r=ET.SubElement(p,q('r'))
    t=ET.SubElement(r,q('t'))
    t.text=new_text
    return r

def fix_frontmatter_titles_and_blanks(root):
    body=root.find(q('body'))
    paras=body.findall('./w:p',NS)
    changed=0
    for i,p in enumerate(paras):
        t=text(p)
        if is_toc_title(t) or safe_heading_level(t)==1:
            break
        if is_cn_abs_title(t):
            set_ppr_basic(p,'center','0')
            set_runs(p,FONT_HEI,'Times New Roman',32,False)
            changed += 1
        elif is_en_abs_title(t):
            replace_para_text(p,'ABSTRACT')
            set_ppr_basic(p,'center','0')
            set_runs(p,'Times New Roman','Times New Roman',32,False)
            changed += 1
        if i>0 and is_en_abs_title(text(paras[i-1])) and text(p)=='' and not p.findall('.//w:drawing',NS):
            set_ppr_basic(p,'center','0')
            set_runs(p,'Times New Roman','Times New Roman',24,False)
            changed += 1
    return changed

def build_static_toc_if_missing(root):
    body=root.find(q('body'))
    nodes=list(body)
    if any(n.tag==q('p') and is_toc_title(text(n)) for n in nodes):
        return 0
    first_h1_idx=None
    entries=[]
    for idx,n in enumerate(nodes):
        if n.tag!=q('p'):
            continue
        t=text(n)
        lvl=safe_heading_level(t)
        if lvl:
            if first_h1_idx is None:
                first_h1_idx=idx
            entries.append((lvl,t,'1'))
    if first_h1_idx is None or not entries:
        return 0
    insert=[]
    title=make_para(chr(0x76ee)+'    '+chr(0x5f55))
    set_ppr_basic(title,'center','0')
    set_runs(title,FONT_HEI,'Times New Roman',32,False)
    insert.append(title)
    for lvl,t,page in entries:
        p=make_para('', f'TOC{min(lvl,3)}')
        replace_para_text(p,t)
        r=ET.SubElement(p,q('r')); ET.SubElement(r,q('tab'))
        r2=ET.SubElement(p,q('r')); tx=ET.SubElement(r2,q('t')); tx.text=page
        set_ppr_basic(p,'left','0')
        ppr=p.find(q('pPr')); ind=ensure(ppr,'ind')
        if lvl==2:
            ind.set(q('left'),'480')
        elif lvl==3:
            ind.set(q('left'),'960')
        set_runs(p,FONT_SONG,'Times New Roman',24,False)
        insert.append(p)
    blank=make_para('')
    set_ppr_basic(blank,'left','0')
    set_runs(blank,FONT_SONG,'Times New Roman',24,False)
    insert.append(blank)
    for off,n in enumerate(insert):
        body.insert(first_h1_idx+off,n)
    return len(entries)

def set_ppr_basic(p, align, first=None, line='440'):
    ppr=ensure(p,'pPr',True)
    jc=ensure(ppr,'jc'); jc.set(q('val'),align)
    ind=ensure(ppr,'ind')
    for a in ('firstLine','left','hanging','right','firstLineChars','leftChars','hangingChars','rightChars'):
        ind.attrib.pop(q(a),None)
    if first is not None: ind.set(q('firstLineChars'),str(first))
    sp=ensure(ppr,'spacing'); sp.attrib.clear(); sp.set(q('before'),'0'); sp.set(q('after'),'0'); sp.set(q('line'),line); sp.set(q('lineRule'),'exact')

def set_runs(p,east,asc,size,bold=False):
    rs=p.findall('./w:r',NS)
    if not rs:
        r=ET.SubElement(p,q('r')); ET.SubElement(r,q('t')).text=''
        rs=[r]
    for r in rs:
        rpr=ensure(r,'rPr',True)
        for c in list(rpr):
            if c.tag in (q('rFonts'),q('sz'),q('szCs'),q('b'),q('bCs'),q('color'),q('u')): rpr.remove(c)
        rf=ET.SubElement(rpr,q('rFonts')); rf.set(q('eastAsia'),east); rf.set(q('ascii'),asc); rf.set(q('hAnsi'),asc); rf.set(q('cs'),asc)
        b=ET.SubElement(rpr,q('b')); b.set(q('val'),'1' if bold else '0')
        bcs=ET.SubElement(rpr,q('bCs')); bcs.set(q('val'),'1' if bold else '0')
        sz=ET.SubElement(rpr,q('sz')); sz.set(q('val'),str(size)); szcs=ET.SubElement(rpr,q('szCs')); szcs.set(q('val'),str(size))
        color=ET.SubElement(rpr,q('color')); color.set(q('val'),'000000')

def fix_safe_headings(root):
    changed=0; restored=0; in_toc=False
    body=root.find(q('body'))
    for p in body.findall('./w:p',NS):
        t=text(p)
        if is_toc_title(t):
            in_toc=True
            continue
        if in_toc:
            # Our static TOC rows end with a visible page-number run. They must never be
            # treated as body headings even if their text starts with "1 ...".
            if re.search(r'\d\s*$', t):
                continue
            if safe_heading_level(t)==1:
                in_toc=False
            else:
                continue
        lvl=safe_heading_level(t); st=style(p)
        if lvl:
            ppr=ensure(p,'pPr',True); ps=ensure(ppr,'pStyle'); ps.set(q('val'),str(2+lvl))
            if lvl==1: set_ppr_basic(p,'center','0'); set_runs(p,FONT_HEI,'Times New Roman',32,False)
            elif lvl==2: set_ppr_basic(p,'left','0'); set_runs(p,FONT_HEI,'Times New Roman',30,False)
            else: set_ppr_basic(p,'left','0'); set_runs(p,FONT_HEI,'Times New Roman',28,False)
            changed+=1
        elif st in ('3','4','5','Heading1','Heading2','Heading3') and t:
            ppr=ensure(p,'pPr',True); ps=ppr.find(q('pStyle'))
            if ps is not None: ppr.remove(ps)
            set_ppr_basic(p,'both','200'); set_runs(p,FONT_SONG,'Times New Roman',24,False)
            restored+=1
    return changed,restored
def fix_toc_styles(styles_root):
    changed=0
    for style in styles_root.findall('.//w:style',NS):
        sid=style.get(q('styleId')) or ''
        if sid not in ('TOC1','TOC2','TOC3','TOC4'): 
            continue
        ppr=ensure(style,'pPr')
        for tag in ('pageBreakBefore','keepNext','keepLines','widowControl'):
            n=ppr.find(q(tag))
            if n is not None:
                ppr.remove(n)
                changed += 1
        jc=ppr.find(q('jc'))
        if jc is None:
            jc=ET.SubElement(ppr,q('jc'))
        jc.set(q('val'),'left')
        ind=ensure(ppr,'ind')
        for a in ('firstLine','left','hanging','right','firstLineChars','leftChars','hangingChars','rightChars'):
            ind.attrib.pop(q(a),None)
        sp=ensure(ppr,'spacing')
        sp.attrib.clear()
        sp.set(q('before'),'0'); sp.set(q('after'),'0'); sp.set(q('line'),'440'); sp.set(q('lineRule'),'exact')
    return changed

def process(src):
    src=Path(src)
    tmp=Path(tempfile.mkdtemp(prefix='finalxml_'))
    try:
        with ZipFile(src) as z: z.extractall(tmp)
        docxml=tmp/'word'/'document.xml'
        tree=ET.parse(str(docxml)); root=tree.getroot()
        fix_sections(root)
        frontmatter_changed = fix_frontmatter_titles_and_blanks(root)
        # TOC is a protected zone: do not create/update/reformat it by default.
        toc_entries = 0
        heading_changed, heading_restored = fix_safe_headings(root)
        tables=fix_tables(root)
        tree.write(str(docxml),xml_declaration=True,encoding='utf-8',standalone='yes')
        stylesxml=tmp/'word'/'styles.xml'
        # TOC styles are also protected by default; keep source/template directory intact.
        toc_changed=0
        repack=src.with_suffix('.tmp.docx')
        if repack.exists():
            repack.unlink()
        with ZipFile(repack,'w',ZIP_DEFLATED) as z:
            for f in tmp.rglob('*'):
                if f.is_file(): z.write(f,f.relative_to(tmp))
        src.unlink(missing_ok=True)
        repack.replace(src)
        print('FINAL_XML_FRONTMATTER_FIXED',frontmatter_changed)
        print('FINAL_XML_TOC_PROTECTED_NO_CREATE',toc_entries)
        print('FINAL_XML_HEADINGS_FIXED',heading_changed)
        print('FINAL_XML_HEADINGS_RESTORED_BODY',heading_restored)
        print('FINAL_XML_TABLES_FIXED',tables)
        print('FINAL_XML_TOC_PROTECTED_STYLE_CHANGES',toc_changed)
    finally:
        shutil.rmtree(tmp,ignore_errors=True)

if __name__=='__main__':
    process(sys.argv[1])
