from zipfile import ZipFile, ZIP_DEFLATED
from pathlib import Path
from lxml import etree as ET
import sys, tempfile, shutil, re

W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS={'w':W.strip('{}')}
TWIPS_PER_CM=567
MARGIN_25=1417
HEADER_16=907
FOOTER_175=992
LINE_22=440
FIRST_2CH=480
H1_TEXTS={'致谢','参考文献','摘  要','ABSTRACT','目  录','摘    要','目    录'}

def q(t): return f'{W}{t}'
def text(p): return ''.join(t.text or '' for t in p.findall('.//w:t',NS)).strip()
def raw_text(p): return ''.join(t.text or '' for t in p.findall('.//w:t',NS))
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
        r=ET.SubElement(p,q('r')); ET.SubElement(r,q('t')).text=' '; rs=[r]
    for r in rs: set_run(r,east,asc,size,bold,color)

def set_single_font_run(r,font='宋体',size=24,bold=False,color='000000'):
    """For blank lines only: one and only one font across eastAsia/ascii/hAnsi/cs."""
    set_run(r,font,font,size,bold,color)

def set_para_mark_single_font(p,font='宋体',size=24,bold=False,color='000000'):
    set_para_mark_font(p,font,font,size,bold,color)

def set_para_mark_font(p,east='宋体',asc='Times New Roman',size=24,bold=False,color='000000'):
    """Set paragraph-mark font too; otherwise blank lines can show as Cambria 11 in WPS/Word."""
    ppr=ensure(p,'pPr',True)
    rpr=ensure(ppr,'rPr')
    for c in list(rpr):
        if c.tag in (q('rFonts'),q('sz'),q('szCs'),q('b'),q('bCs'),q('color'),q('u')): rpr.remove(c)
    rf=ET.SubElement(rpr,q('rFonts')); rf.set(q('eastAsia'),east); rf.set(q('ascii'),asc); rf.set(q('hAnsi'),asc); rf.set(q('cs'),asc)
    b=ET.SubElement(rpr,q('b')); b.set(q('val'),'1' if bold else '0')
    bcs=ET.SubElement(rpr,q('bCs')); bcs.set(q('val'),'1' if bold else '0')
    sz=ET.SubElement(rpr,q('sz')); sz.set(q('val'),str(size))
    szcs=ET.SubElement(rpr,q('szCs')); szcs.set(q('val'),str(size))
    c=ET.SubElement(rpr,q('color')); c.set(q('val'),color)


def set_ppr(p,align='both',first_chars=200,line=LINE_22,before=0,after=0):
    ppr=ensure(p,'pPr',True)
    jc=ensure(ppr,'jc'); jc.set(q('val'),align)
    sp=ensure(ppr,'spacing'); sp.set(q('before'),str(before)); sp.set(q('after'),str(after)); sp.set(q('line'),str(line)); sp.set(q('lineRule'),'exact')
    ind=ensure(ppr,'ind')
    for a in ('firstLine','left','hanging','right','firstLineChars','leftChars','hangingChars','rightChars'):
        ind.attrib.pop(q(a),None)
    # Strict 2-character indent: use firstLineChars=200, not twips/cm conversion.
    if first_chars is not None:
        ind.set(q('firstLineChars'),str(first_chars))
        # Some WPS/Word styles reintroduce firstLine=480 via style inheritance. Explicitly
        # clear it again after setting firstLineChars so strict 2-character indent wins.
        ind.attrib.pop(q('firstLine'),None)
    set_para_mark_font(p,'宋体','Times New Roman',24,False,'000000')

def format_blank(p):
    ppr=ensure(p,'pPr',True)
    ps=ppr.find(q('pStyle'))
    if ps is not None: ppr.remove(ps)
    set_ppr(p,'left',0,LINE_22,0,0)
    for r in list(p.findall('./w:r',NS)): p.remove(r)
    r=ET.SubElement(p,q('r')); set_single_font_run(r,'宋体',24,False,'000000')
    set_para_mark_single_font(p,'宋体',24,False,'000000')
    t=ET.SubElement(r,q('t')); t.text=' '; t.set('{http://www.w3.org/XML/1998/namespace}space','preserve')

def set_blank_font(p,east='宋体',asc='Times New Roman',size=24):
    ppr=ensure(p,'pPr',True)
    ps=ppr.find(q('pStyle'))
    if ps is not None: ppr.remove(ps)
    set_ppr(p,'left',0,LINE_22,0,0)
    for r in list(p.findall('./w:r',NS)): p.remove(r)
    # Blank lines must have exactly one font. Use east as the single requested font;
    # normal blanks pass 宋体, the only exception passes Times New Roman.
    font=east
    r=ET.SubElement(p,q('r')); set_single_font_run(r,font,size,False,'000000')
    set_para_mark_single_font(p,font,size,False,'000000')
    t=ET.SubElement(r,q('t')); t.text=' '; t.set('{http://www.w3.org/XML/1998/namespace}space','preserve')


def fix_english_abstract_title_and_blanks(root):
    paras=root.findall('.//w:p',NS)
    fixed=0
    for i,p in enumerate(paras):
        if text(p) == 'ABSTRACT':
            # ABSTRACT title itself: Times New Roman, 小三/16pt equivalent already used in earlier rules.
            set_para_mark_font(p,'Times New Roman','Times New Roman',32,True,'000000')
            set_all_runs(p,'Times New Roman','Times New Roman',32,True,'000000')
            # Only the blank line BELOW ABSTRACT is Times New Roman.
            # The blank line ABOVE ABSTRACT, and all other H1 surrounding blanks, stay 宋体.
            j = i + 1
            if 0 <= j < len(paras) and (text(paras[j]) == '' or raw_text(paras[j]).strip() == ''):
                set_blank_font(paras[j],'Times New Roman','Times New Roman',24)
                fixed += 1
            fixed += 1
    return fixed


def set_all_section_page_setup(root):
    changed=0
    for sect in root.findall('.//w:sectPr',NS):
        pgMar=ensure(sect,'pgMar')
        pgMar.set(q('top'),str(MARGIN_25)); pgMar.set(q('bottom'),str(MARGIN_25))
        pgMar.set(q('left'),str(MARGIN_25)); pgMar.set(q('right'),str(MARGIN_25))
        pgMar.set(q('header'),str(HEADER_16)); pgMar.set(q('footer'),str(FOOTER_175))
        # Keep gutter at 0 unless explicitly present.
        pgMar.set(q('gutter'),'0')
        changed+=1
    return changed

def is_heading(p): return style(p) in ('3','4','5','6','Heading1','Heading2','Heading3','Heading4')
def is_toc_line(t): return '\t' in raw_text_cur or bool(re.match(r'^(\d+(?:\.\d+)*\s+.+|致谢|参考文献)\s*\d+$',t))
def is_caption_or_note(t): return bool(re.match(r'^[图表]\d+\.\d+\s+',t)) or t.startswith('注：')
def is_reference(t): return bool(re.match(r'^\[\d+\]\s+',t))
def has_drawing(p): return bool(p.findall('.//w:drawing',NS))
def looks_like_body_sentence(t):
    # Headings may contain ：、（）/ etc. Only reject obvious long sentences.
    return len(t) > 70 or bool(re.search(r'[。；;]$', t))

def heading_level_by_text(t):
    if looks_like_body_sentence(t): return None
    if re.match(r'^\d+\.\d+\.\d+\s*\S.{0,70}$', t): return 3
    if re.match(r'^\d+\.\d+\s*\S.{0,65}$', t): return 2
    if t in H1_TEXTS or re.match(r'^\d+\s+\S.{0,55}$', t): return 1
    return None

def is_front_cover_line(t):
    return any(k in t for k in ('编号', '审定等级', '论文（设计）题目', '学院：', '专业：', '班级：', '姓名：', '学号：', '指导教师', '答辩组负责人', '填表时间', '重庆移通学院教务处制'))

def fix_paragraphs(root):
    paras=root.findall('.//w:p',NS)
    state='front'
    blanks=body_indent=statement=abstract=0
    for p in paras:
        t=text(p)
        st=style(p)
        lvl=heading_level_by_text(t)
        if t=='原创性声明':
            state='statement_title'
            continue
        if t in ('摘    要','摘  要'):
            state='abstract_cn'
            continue
        if t=='ABSTRACT':
            state='abstract_en'
            continue
        if t in ('目    录','目  录'):
            state='toc'
            continue
        if lvl==1 and t not in ('摘    要','摘  要','ABSTRACT','目    录','目  录'):
            state='body'
        if t=='参考文献':
            state='refs'
        if has_drawing(p):
            # Never rewrite runs in picture paragraphs; that deletes w:drawing objects.
            continue
        if t=='' or raw_text(p).strip()=='' or raw_text(p)==' ':
            format_blank(p); blanks+=1; continue
        if is_heading(p):
            continue
        if is_caption_or_note(t) or is_reference(t):
            continue
        if state=='statement_title':
            # 原创性声明标题后的正文：两端对齐，首行缩进两字符。签名/日期行不强改。
            if '本人签名' in t or re.match(r'^年\s*月\s*日$',t):
                continue
            set_ppr(p,'both',200,LINE_22,0,0); set_all_runs(p,'宋体','Times New Roman',24,False,'000000')
            statement+=1; continue
        if state in ('abstract_cn','abstract_en'):
            if t.startswith(('关键词','Key words','Keywords','Key word')):
                continue
            set_ppr(p,'both',200,LINE_22,0,0); set_all_runs(p,'宋体' if state=='abstract_cn' else 'Times New Roman','Times New Roman',24,False,'000000')
            abstract+=1; continue
        if state=='body':
            # 普通正文、致谢正文等需要首行缩进两字符。
            if not is_front_cover_line(t):
                set_ppr(p,'both',200,LINE_22,0,0); set_all_runs(p,'宋体','Times New Roman',24,False,'000000')
                body_indent+=1
    # Final body fallback: once the first real H1 body chapter starts, force all plain
    # body paragraphs until references to strict two-character indent.
    in_body=False
    for p in paras:
        t=text(p)
        lvl=heading_level_by_text(t)
        if lvl==1 and t not in ('摘    要','摘  要','ABSTRACT','目    录','目  录'):
            in_body=True
            if t=='参考文献':
                in_body=False
            continue
        if t=='参考文献':
            in_body=False
            continue
        if not in_body:
            continue
        if has_drawing(p) or not t or is_heading(p) or is_caption_or_note(t) or is_reference(t):
            continue
        set_ppr(p,'both',200,LINE_22,0,0); set_all_runs(p,'宋体','Times New Roman',24,False,'000000')
    return blanks,body_indent,statement,abstract

def process(src,out):
    src=Path(src); out=Path(out); out.parent.mkdir(parents=True,exist_ok=True)
    tmp=Path(tempfile.mkdtemp(prefix='step20base_'))
    try:
        with ZipFile(src) as z: z.extractall(tmp)
        docxml=tmp/'word'/'document.xml'; tree=ET.parse(str(docxml)); root=tree.getroot()
        sections=set_all_section_page_setup(root)
        blanks,body_indent,statement,abstract=fix_paragraphs(root)
        abstract_blank=fix_english_abstract_title_and_blanks(root)
        tree.write(str(docxml),xml_declaration=True,encoding='utf-8',standalone='yes')
        repack=out.with_suffix('.tmp.docx')
        with ZipFile(repack,'w',ZIP_DEFLATED) as z:
            for f in tmp.rglob('*'):
                if f.is_file(): z.write(f,f.relative_to(tmp))
        shutil.move(str(repack),str(out))
        print('OUTPUT:',out)
        print('SECTIONS_PAGE_SETUP_FIXED:',sections)
        print('BLANK_PARAS_FIXED:',blanks)
        print('BODY_INDENT_2CH_FIXED:',body_indent)
        print('STATEMENT_JUSTIFY_INDENT_FIXED:',statement)
        print('ABSTRACT_INDENT_FIXED:',abstract)
        print('EN_ABSTRACT_TITLE_BLANKS_FIXED:',abstract_blank)
    finally:
        shutil.rmtree(tmp,ignore_errors=True)
if __name__=='__main__':
    if len(sys.argv)!=3: raise SystemExit('usage: python step20_xml_fix_baseline_layout.py in.docx out.docx')
    process(sys.argv[1],sys.argv[2])
