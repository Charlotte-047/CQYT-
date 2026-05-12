from __future__ import annotations
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree as ET
import os, re, shutil, sys, tempfile

W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS={'w':W}
def q(t): return f'{{{W}}}{t}'
def norm(t): return re.sub(r'\s+','',t or '')
def text(p): return ''.join(t.text or '' for t in p.findall('.//w:t',NS)).strip()
def raw_text(p): return ''.join(t.text or '' for t in p.findall('.//w:t',NS))
def ensure(parent, tag, first=False):
    n=parent.find(q(tag))
    if n is None:
        n=ET.Element(q(tag))
        if first: parent.insert(0,n)
        else: parent.append(n)
    return n

def is_toc_title(t): return norm(t)==chr(0x76ee)+chr(0x5f55)
def is_real_body_h1(p):
    t=text(p)
    ppr=p.find(q('pPr')); ps=None if ppr is None else ppr.find(q('pStyle'))
    st='' if ps is None else ps.get(q('val')) or ''
    # Only the real formatted body H1 style ends the TOC. TOC entries may look
    # like "1 绪论1" but use TOC styles (26/31/17), so never stop on text alone.
    return st in ('2','Heading1') and bool(re.match(r'^\d+\s+\S+',t))

def set_para_line_22(p, align=None, first_chars=None):
    ppr=ensure(p,'pPr',True)
    sp=ppr.find(q('spacing'))
    if sp is None: sp=ET.SubElement(ppr,q('spacing'))
    sp.set(q('line'),'440')
    sp.set(q('lineRule'),'exact')
    sp.set(q('before'),'0')
    sp.set(q('after'),'0')
    if align is not None:
        jc=ppr.find(q('jc'))
        if jc is None: jc=ET.SubElement(ppr,q('jc'))
        jc.set(q('val'),align)
    ind=ppr.find(q('ind'))
    if ind is None: ind=ET.SubElement(ppr,q('ind'))
    if first_chars is None:
        for k in ('firstLine','firstLineChars','hanging','hangingChars'):
            ind.attrib.pop(q(k),None)
    else:
        ind.set(q('firstLineChars'),str(first_chars))
        ind.attrib.pop(q('firstLine'),None)
        ind.attrib.pop(q('hanging'),None)
        ind.attrib.pop(q('hangingChars'),None)

def set_run_font_all(p, font='宋体', size=None, bold=None):
    for r in p.findall('.//w:r',NS):
        rpr=ensure(r,'rPr',True)
        rf=rpr.find(q('rFonts'))
        if rf is None: rf=ET.SubElement(rpr,q('rFonts'))
        for a in ('ascii','hAnsi','eastAsia','cs'):
            rf.set(q(a),font)
        if size is not None:
            sz=rpr.find(q('sz'))
            if sz is None: sz=ET.SubElement(rpr,q('sz'))
            sz.set(q('val'),str(size))
            szcs=rpr.find(q('szCs'))
            if szcs is None: szcs=ET.SubElement(rpr,q('szCs'))
            szcs.set(q('val'),str(size))
        if bold is not None:
            for tag in ('b','bCs'):
                b=rpr.find(q(tag))
                if b is None: b=ET.SubElement(rpr,q(tag))
                b.set(q('val'),'1' if bold else '0')
    # Paragraph mark font matters for blank/TOC lines too.
    ppr=ensure(p,'pPr',True); rpr=ensure(ppr,'rPr')
    rf=rpr.find(q('rFonts'))
    if rf is None: rf=ET.SubElement(rpr,q('rFonts'))
    for a in ('ascii','hAnsi','eastAsia','cs'):
        rf.set(q(a),font)
    if size is not None:
        sz=rpr.find(q('sz'))
        if sz is None: sz=ET.SubElement(rpr,q('sz'))
        sz.set(q('val'),str(size))
        szcs=rpr.find(q('szCs'))
        if szcs is None: szcs=ET.SubElement(rpr,q('szCs'))
        szcs.set(q('val'),str(size))

def set_run_font_mixed(p, east='宋体', latin='Times New Roman', size=None, bold=None):
    for r in p.findall('.//w:r',NS):
        rpr=ensure(r,'rPr',True)
        rf=rpr.find(q('rFonts'))
        if rf is None: rf=ET.SubElement(rpr,q('rFonts'))
        rf.set(q('eastAsia'),east)
        rf.set(q('ascii'),latin)
        rf.set(q('hAnsi'),latin)
        rf.set(q('cs'),latin)
        if size is not None:
            sz=rpr.find(q('sz'))
            if sz is None: sz=ET.SubElement(rpr,q('sz'))
            sz.set(q('val'),str(size))
            szcs=rpr.find(q('szCs'))
            if szcs is None: szcs=ET.SubElement(rpr,q('szCs'))
            szcs.set(q('val'),str(size))
        if bold is not None:
            for tag in ('b','bCs'):
                b=rpr.find(q(tag))
                if b is None: b=ET.SubElement(rpr,q(tag))
                b.set(q('val'),'1' if bold else '0')
    ppr=ensure(p,'pPr',True); rpr=ensure(ppr,'rPr')
    rf=rpr.find(q('rFonts'))
    if rf is None: rf=ET.SubElement(rpr,q('rFonts'))
    rf.set(q('eastAsia'),east); rf.set(q('ascii'),latin); rf.set(q('hAnsi'),latin); rf.set(q('cs'),latin)
    if size is not None:
        sz=rpr.find(q('sz'))
        if sz is None: sz=ET.SubElement(rpr,q('sz'))
        sz.set(q('val'),str(size))
        szcs=rpr.find(q('szCs'))
        if szcs is None: szcs=ET.SubElement(rpr,q('szCs'))
        szcs.set(q('val'),str(size))

def force_toc_songti_22(root):
    body=root.find('.//w:body',NS); children=list(body)
    in_toc=False; changed=0
    for n in children:
        if n.tag!=q('p'): continue
        t=text(n)
        if is_toc_title(t):
            in_toc=True
            set_para_line_22(n,'center',None)
            set_run_font_all(n,'宋体',32,False)
            changed+=1
            continue
        if in_toc:
            if is_real_body_h1(n):
                in_toc=False
                continue
            # All TOC lines, blank TOC lines, and field/hyperlink runs: Songti + fixed 22.
            set_para_line_22(n,None,None)
            set_run_font_all(n,'宋体',24,False)
            changed+=1
    return changed

def clear_children(parent, tag):
    for n in list(parent.findall(q(tag))):
        parent.remove(n)

def set_border(parent, edge, val='single', sz='12', color='auto'):
    n=parent.find(q(edge))
    if n is None: n=ET.SubElement(parent,q(edge))
    n.set(q('val'),val); n.set(q('sz'),str(sz)); n.set(q('space'),'0'); n.set(q('color'),color)
    return n

def nil_border(parent, edge):
    n=parent.find(q(edge))
    if n is None: n=ET.SubElement(parent,q(edge))
    n.set(q('val'),'nil'); n.set(q('sz'),'0'); n.set(q('space'),'0'); n.set(q('color'),'auto')
    return n

def set_cell_text_format(p):
    set_para_line_22(p,'center',0)
    # table body uses single/auto line spacing per earlier rules; keep 22 only for captions/TOC.
    ppr=ensure(p,'pPr',True)
    sp=ppr.find(q('spacing'))
    if sp is None: sp=ET.SubElement(ppr,q('spacing'))
    sp.set(q('line'),'240'); sp.set(q('lineRule'),'auto'); sp.set(q('before'),'0'); sp.set(q('after'),'0')
    set_run_font_mixed(p,'宋体','Times New Roman',21,False)
    if ppr.find(q('keepLines')) is None: ET.SubElement(ppr,q('keepLines'))

def force_three_line_tables(root):
    """Force tables to match t0 in 三线表标准格式.docx.

    Standard t0 draws visible lines on cell borders, not tblBorders:
    - header row cell top: 1.5pt (w:sz=12)
    - header row cell bottom: 0.75pt (w:sz=6)
    - last row cell bottom: 1.5pt (w:sz=12)
    - no left/right/internal/body-row borders
    """
    tables=root.findall('.//w:tbl',NS); changed=0
    for tbl in tables:
        tblPr=ensure(tbl,'tblPr',True)
        tb=tblPr.find(q('tblBorders'))
        if tb is None: tb=ET.SubElement(tblPr,q('tblBorders'))
        for e in ('top','left','bottom','right','insideH','insideV'):
            nil_border(tb,e)
        rows=tbl.findall('./w:tr',NS)
        for ri,row in enumerate(rows):
            trPr=ensure(row,'trPr',True)
            if trPr.find(q('cantSplit')) is None: ET.SubElement(trPr,q('cantSplit'))
            for cell in row.findall('./w:tc',NS):
                tcPr=ensure(cell,'tcPr',True)
                cb=tcPr.find(q('tcBorders'))
                if cb is None: cb=ET.SubElement(tcPr,q('tcBorders'))
                for e in ('top','left','bottom','right','insideH','insideV','tl2br','tr2bl'):
                    nil_border(cb,e)
                if ri==0:
                    set_border(cb,'top','single','12')
                    set_border(cb,'bottom','single','6')
                if ri==len(rows)-1:
                    set_border(cb,'bottom','single','12')
                for p in cell.findall('.//w:p',NS):
                    set_cell_text_format(p)
                    ppr=ensure(p,'pPr',True)
                    kn=ppr.find(q('keepNext'))
                    if ri < len(rows)-1:
                        if kn is None: ET.SubElement(ppr,q('keepNext'))
                    elif kn is not None:
                        ppr.remove(kn)
        changed+=1
    return changed

def process(src,out=None):
    src=Path(src); out=Path(out) if out else src
    with tempfile.TemporaryDirectory() as td:
        td=Path(td)
        with ZipFile(src,'r') as z: z.extractall(td)
        doc=td/'word'/'document.xml'
        tree=ET.parse(str(doc)); root=tree.getroot()
        toc=force_toc_songti_22(root)
        tbl=force_three_line_tables(root)
        tree.write(str(doc),xml_declaration=True,encoding='utf-8',standalone='yes')
        tmp=out.with_suffix(out.suffix+'.tmp_strict')
        with ZipFile(tmp,'w',ZIP_DEFLATED) as z:
            for folder,_,files in os.walk(td):
                for f in files:
                    p=Path(folder)/f
                    z.write(p,p.relative_to(td).as_posix())
        shutil.move(str(tmp),str(out))
    print('FINAL_STRICT_TOC_PARAS_FIXED',toc)
    print('FINAL_STRICT_TABLES_FIXED',tbl)

if __name__=='__main__':
    if len(sys.argv)==2: process(sys.argv[1])
    elif len(sys.argv)==3: process(sys.argv[1],sys.argv[2])
    else:
        print('usage: final_strict_toc_tables_v1.py input.docx [output.docx]')
        raise SystemExit(2)


