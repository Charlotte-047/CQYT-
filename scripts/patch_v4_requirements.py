from zipfile import ZipFile, ZIP_DEFLATED
from pathlib import Path
from lxml import etree as ET
import sys, tempfile, shutil, re

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
REL = '{http://schemas.openxmlformats.org/package/2006/relationships}'
NS = {'w': W.strip('{}'), 'r': R.strip('{}'), 'rel': REL.strip('{}')}


def q(t): return f'{W}{t}'
def text(p): return ''.join(t.text or '' for t in p.findall('.//w:t', NS)).strip()
def style(p):
    ppr = p.find(q('pPr'))
    ps = None if ppr is None else ppr.find(q('pStyle'))
    return '' if ps is None else ps.get(q('val')) or ''

def ensure(parent, tag, ns=W, first=False):
    node = parent.find(f'{ns}{tag}')
    if node is None:
        node = ET.Element(f'{ns}{tag}')
        if first: parent.insert(0, node)
        else: parent.append(node)
    return node

def remove_null_rels(tmp):
    for relp in tmp.rglob('*.rels'):
        tree = ET.parse(str(relp)); root = tree.getroot(); removed = 0
        for rel in list(root):
            if rel.get('Target') in ('../NULL', 'NULL', ''):
                root.remove(rel); removed += 1
        if removed:
            tree.write(str(relp), xml_declaration=True, encoding='utf-8', standalone='yes')

def set_run(r, east='宋体', asc='Times New Roman', size=24, bold=False):
    rpr = ensure(r, 'rPr', first=True)
    for child in list(rpr):
        if child.tag in (q('rFonts'), q('sz'), q('szCs'), q('b'), q('bCs'), q('color'), q('u')):
            rpr.remove(child)
    rf = ET.SubElement(rpr, q('rFonts'))
    rf.set(q('eastAsia'), east); rf.set(q('ascii'), asc); rf.set(q('hAnsi'), asc); rf.set(q('cs'), asc)
    b = ET.SubElement(rpr, q('b')); b.set(q('val'), '1' if bold else '0')
    bcs = ET.SubElement(rpr, q('bCs')); bcs.set(q('val'), '1' if bold else '0')
    sz = ET.SubElement(rpr, q('sz')); sz.set(q('val'), str(size))
    szcs = ET.SubElement(rpr, q('szCs')); szcs.set(q('val'), str(size))

def ensure_run(p):
    rs = p.findall('./w:r', NS)
    if rs: return rs
    r = ET.SubElement(p, q('r')); ET.SubElement(r, q('t')).text = ''
    return [r]

def set_all_runs(p, east='宋体', asc='Times New Roman', size=24, bold=False):
    for r in ensure_run(p):
        set_run(r, east, asc, size, bold)

def set_ppr(p, align=None, first=None, left=None, line=440, before=0, after=0):
    ppr = ensure(p, 'pPr', first=True)
    if align is not None:
        jc = ensure(ppr, 'jc'); jc.set(q('val'), align)
    sp = ensure(ppr, 'spacing')
    sp.set(q('before'), str(before)); sp.set(q('after'), str(after)); sp.set(q('line'), str(line)); sp.set(q('lineRule'), 'exact')
    ind = ensure(ppr, 'ind')
    for a in ('firstLine', 'left', 'hanging', 'right', 'firstLineChars', 'leftChars', 'hangingChars', 'rightChars'):
        ind.attrib.pop(q(a), None)
    if first is not None: ind.set(q('firstLine'), str(first))
    if left is not None: ind.set(q('left'), str(left))

def make_blank():
    p = ET.Element(q('p')); set_ppr(p, 'left', 0, None, 440, 0, 0); set_all_runs(p, '宋体', 'Times New Roman', 24, False); return p

def add_page_break_to_para(p):
    r = ET.SubElement(p, q('r')); br = ET.SubElement(r, q('br')); br.set(q('type'), 'page')

def paragraph_has_page_break(p): return bool(p.findall('.//w:br', NS))
def is_blank(p): return text(p) == ''

def replace_text(p, new):
    for r in list(p.findall('./w:r', NS)): p.remove(r)
    r = ET.SubElement(p, q('r')); t = ET.SubElement(r, q('t')); t.text = new
    return r

def looks_like_body_sentence(t):
    # Headings may contain ：、（）/ etc. Only reject obvious long sentences.
    return len(t) > 70 or bool(re.search(r'[。；;]$', t))

def is_h1_text(t):
    if t in ('致谢', '参考文献'):
        return True
    if looks_like_body_sentence(t):
        return False
    return bool(re.match(r'^\d+\s+\S.{0,55}$', t))

def is_h2_text(t):
    if looks_like_body_sentence(t):
        return False
    return bool(re.match(r'^\d+\.\d+\s*\S.{0,65}$', t))

def is_h3_text(t):
    if looks_like_body_sentence(t):
        return False
    return bool(re.match(r'^\d+\.\d+\.\d+\s*\S.{0,70}$', t))

def is_toc_like_text(t): return bool(re.match(r'^(\d+(?:\.\d+)*\s+.+|致谢|参考文献)\s*\d+$', t))

def format_abstract_body(p, english=False):
    set_ppr(p, 'both', 480, None, 440, 0, 0)
    set_all_runs(p, 'Times New Roman' if english else '宋体', 'Times New Roman', 24, False)

def format_reference_para(p):
    set_ppr(p, 'left', 0, 0, 440, 0, 0)
    ppr = p.find(q('pPr')); ind = ensure(ppr, 'ind'); ind.set(q('hanging'), '480')
    set_all_runs(p, '宋体', 'Times New Roman', 24, False)

def patch_document(tmp):
    docxml = tmp / 'word' / 'document.xml'; tree = ET.parse(str(docxml)); root = tree.getroot(); body = root.find(f'{W}body')
    paras = body.findall('./w:p', NS)
    abstract = False; enabs = False; toc = False; real_body = False; refs = False
    for p in paras:
        t = text(p); st = style(p)
        if t == '原创性声明':
            abstract = enabs = toc = False
            set_ppr(p, 'center', 0, None); set_all_runs(p, '宋体', 'Times New Roman', 44, True); continue
        if t in ('摘    要', '摘  要'):
            abstract = True; enabs = False; toc = False
            set_ppr(p, 'center', 0, None); set_all_runs(p, '黑体', 'Times New Roman', 32, False); continue
        if t == 'ABSTRACT':
            abstract = False; enabs = True; toc = False
            set_ppr(p, 'center', 0, None); set_all_runs(p, 'Times New Roman', 'Times New Roman', 32, True); continue
        if t in ('目    录', '目  录'):
            abstract = False; enabs = False; toc = True
            set_ppr(p, 'center', 0, None); set_all_runs(p, '黑体', 'Times New Roman', 32, False); continue
        # Real body starts only on a true visible H1 title, never on TOC lines like '1 绪论........1'.
        if (t == '1 绪论' or is_h1_text(t)) and not toc and not is_toc_like_text(t):
            abstract = enabs = toc = False; real_body = True
        if real_body and t == '参考文献': refs = True
        if abstract:
            if t.startswith('关键词：'):
                set_ppr(p, 'left', 0, 0)
            elif t:
                format_abstract_body(p, False)
        if enabs:
            if t.startswith(('Key words:', 'Keywords:', 'Key words：', 'Key word:', 'Key word：')):
                set_ppr(p, 'left', 0, 0)
            elif t:
                format_abstract_body(p, True)
        if toc:
            if p.findall('.//w:tab', NS):
                set_ppr(p, 'left', 0, 0, 440, 0, 0)
                set_all_runs(p, '宋体', 'Times New Roman', 24, False)
            continue
        if real_body:
            if is_h1_text(t):
                ppr = ensure(p, 'pPr', True); ps = ensure(ppr, 'pStyle', True); ps.set(q('val'), 'Heading1')
                set_ppr(p, 'center', 0, None); set_all_runs(p, '黑体', 'Times New Roman', 32, False)
            elif is_h2_text(t):
                ppr = ensure(p, 'pPr', True); ps = ensure(ppr, 'pStyle', True); ps.set(q('val'), 'Heading2')
                set_ppr(p, 'left', 0, None); set_all_runs(p, '黑体', 'Times New Roman', 30, False)
            elif is_h3_text(t):
                ppr = ensure(p, 'pPr', True); ps = ensure(ppr, 'pStyle', True); ps.set(q('val'), 'Heading3')
                set_ppr(p, 'left', 0, None); set_all_runs(p, '黑体', 'Times New Roman', 28, False)
            elif refs and t and re.match(r'^\[?\d+\]?', t):
                format_reference_para(p)
            elif t:
                set_ppr(p, 'both', 480, None); set_all_runs(p, '宋体', 'Times New Roman', 24, False)
            else:
                set_ppr(p, 'left', 0, None); set_all_runs(p, '宋体', 'Times New Roman', 24, False)
    children = list(body); i = 0; real_body = False
    while i < len(children):
        node = children[i]
        if node.tag == q('p'):
            t = text(node); st = style(node)
            if is_h1_text(t):
                real_body = True
            if real_body and is_h1_text(t):
                idx = list(body).index(node)
                if idx > 0:
                    prev = list(body)[idx - 1]
                    if prev.tag == q('p') and not paragraph_has_page_break(prev): add_page_break_to_para(prev)
                children = list(body); idx = list(body).index(node)
                if idx == 0 or list(body)[idx - 1].tag != q('p') or not is_blank(list(body)[idx - 1]):
                    body.insert(idx, make_blank()); i += 1; node = list(body)[idx + 1]
                children = list(body); idx = list(body).index(node)
                if idx + 1 >= len(children) or children[idx + 1].tag != q('p') or not is_blank(children[idx + 1]):
                    body.insert(idx + 1, make_blank())
        i += 1; children = list(body)
    tree.write(str(docxml), xml_declaration=True, encoding='utf-8', standalone='yes')

def patch_headers_footers(tmp):
    docxml = tmp / 'word' / 'document.xml'; tree = ET.parse(str(docxml)); root = tree.getroot(); body = root.find(f'{W}body')
    sects = body.findall('.//w:sectPr', NS)
    if not sects:
        sect = ET.SubElement(body, q('sectPr')); sects = [sect]
    header_path = tmp / 'word' / 'header_autogen.xml'; footer_path = tmp / 'word' / 'footer_autogen.xml'
    header_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="{W.strip('{}')}" xmlns:r="{R.strip('{}')}"><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:rFonts w:eastAsia="宋体" w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t>毕业论文（设计）</w:t></w:r></w:p></w:hdr>'''
    footer_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="{W.strip('{}')}" xmlns:r="{R.strip('{}')}"><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:rFonts w:eastAsia="宋体" w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>'''
    header_path.write_text(header_xml, encoding='utf-8'); footer_path.write_text(footer_xml, encoding='utf-8')
    relp = tmp / 'word' / '_rels' / 'document.xml.rels'; rtree = ET.parse(str(relp)); rroot = rtree.getroot()
    ids = [int((rel.get('Id') or 'rId0').replace('rId', '')) for rel in rroot if (rel.get('Id') or '').startswith('rId') and (rel.get('Id') or 'rId0')[3:].isdigit()]
    hid = f'rId{max(ids + [0]) + 1}'; fid = f'rId{max(ids + [0]) + 2}'
    ET.SubElement(rroot, f'{REL}Relationship', Id=hid, Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/header', Target='header_autogen.xml')
    ET.SubElement(rroot, f'{REL}Relationship', Id=fid, Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer', Target='footer_autogen.xml')
    rtree.write(str(relp), xml_declaration=True, encoding='utf-8', standalone='yes')
    for sect in sects:
        for old in sect.findall(q('headerReference')) + sect.findall(q('footerReference')): sect.remove(old)
        hr = ET.Element(q('headerReference')); hr.set(q('type'), 'default'); hr.set(f'{R}id', hid); sect.insert(0, hr)
        fr = ET.Element(q('footerReference')); fr.set(q('type'), 'default'); fr.set(f'{R}id', fid); sect.insert(1, fr)
    tree.write(str(docxml), xml_declaration=True, encoding='utf-8', standalone='yes')

def process(src, out):
    src = Path(src); out = Path(out); out.parent.mkdir(parents=True, exist_ok=True)
    tmp = Path(tempfile.mkdtemp(prefix='patchv4_'))
    try:
        with ZipFile(src) as z: z.extractall(tmp)
        remove_null_rels(tmp)
        patch_document(tmp)
        patch_headers_footers(tmp)
        for xml in (tmp / 'word').rglob('*.xml'):
            raw = xml.read_text(encoding='utf-8', errors='ignore')
            for bad in ('MS 明朝', 'MS Mincho', 'MS 明朗'): raw = raw.replace(bad, '宋体')
            xml.write_text(raw, encoding='utf-8')
        repack = out.with_suffix('.tmp.docx')
        with ZipFile(repack, 'w', ZIP_DEFLATED) as z:
            for f in tmp.rglob('*'):
                if f.is_file(): z.write(f, f.relative_to(tmp))
        shutil.move(str(repack), str(out))
        print('OUTPUT:', out)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

if __name__ == '__main__':
    if len(sys.argv) != 3: raise SystemExit('usage: python patch_v4_requirements.py in.docx out.docx')
    process(sys.argv[1], sys.argv[2])
