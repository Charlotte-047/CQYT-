from zipfile import ZipFile, ZIP_DEFLATED
from pathlib import Path
from lxml import etree as ET
import sys, tempfile, shutil, re

W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
WP='{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
NS={'w':W.strip('{}'),'wp':WP.strip('{}'),'r':R.strip('{}')}

def q(t): return f'{W}{t}'
def wp(t): return f'{WP}{t}'
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

def set_ppr(p,align='both',first='480',line='440',exact=True,keep_next=False,keep_together=False):
    ppr=ensure(p,'pPr',True)
    # remove stale heading style for normal media blanks/body
    jc=ensure(ppr,'jc'); jc.set(q('val'),align)
    sp=ensure(ppr,'spacing'); sp.set(q('before'),'0'); sp.set(q('after'),'0'); sp.set(q('line'),line); sp.set(q('lineRule'),'exact' if exact else 'auto')
    ind=ensure(ppr,'ind')
    for a in ('firstLine','left','hanging','right','firstLineChars','leftChars','hangingChars','rightChars'):
        ind.attrib.pop(q(a),None)
    if first is not None: ind.set(q('firstLine'),first)
    for tag,want in (('keepNext',keep_next),('keepLines',keep_together)):
        old=ppr.find(q(tag))
        if want and old is None: ET.SubElement(ppr,q(tag))
        if not want and old is not None: ppr.remove(old)

def clear_pstyle(p):
    ppr=ensure(p,'pPr',True); ps=ppr.find(q('pStyle'))
    if ps is not None: ppr.remove(ps)

def make_body_para(txt):
    p=ET.Element(q('p')); clear_pstyle(p); set_ppr(p,'both','480','440',True)
    r=ET.SubElement(p,q('r')); set_run(r,'宋体','Times New Roman',24,False,'000000')
    t=ET.SubElement(r,q('t')); t.text=txt
    return p

def make_blank():
    p=ET.Element(q('p')); clear_pstyle(p); set_ppr(p,'left','0','440',True)
    r=ET.SubElement(p,q('r')); set_run(r,'宋体','Times New Roman',24,False,'000000'); ET.SubElement(r,q('t')).text=''
    return p

def make_caption(txt):
    p=ET.Element(q('p')); clear_pstyle(p); set_ppr(p,'center','0','440',True,True,True)
    r=ET.SubElement(p,q('r')); set_run(r,'宋体','Times New Roman',21,False,'000000')
    ET.SubElement(r,q('t')).text=txt
    return p

def is_blank(n): return n is not None and n.tag==q('p') and text(n)==''
def has_drawing(p): return bool(p.findall('.//w:drawing',NS))
def is_h1(p):
    if p.tag!=q('p'): return False
    t=text(p)
    # Current template: style 2 is real Heading 1. Also accept text-detected H1
    # so media/table repair still works if style mapping is imperfect.
    return style(p) in ('2','Heading1') or bool(re.match(r'^\d+\s+\S+',t)) or t in ('致谢','参考文献')
def h1_chapter(p):
    m=re.match(r'^(\d+)\s+',text(p)); return int(m.group(1)) if m else None

def convert_anchor_to_inline_in_para(p):
    count=0
    for anchor in list(p.findall('.//wp:anchor',NS)):
        inline=ET.Element(wp('inline'))
        for child in list(anchor):
            if child.tag in (wp('simplePos'),wp('positionH'),wp('positionV'),wp('wrapSquare'),wp('wrapTight'),wp('wrapThrough'),wp('wrapTopAndBottom'),wp('wrapNone'),wp('wrapBehindDoc')):
                continue
            inline.append(child)
        inline.set('distT','0'); inline.set('distB','0'); inline.set('distL','0'); inline.set('distR','0')
        anchor.getparent().replace(anchor,inline); count+=1
    return count

def drawing_runs_deep(p):
    # collect top-level runs containing drawings; if drawing is nested, clone only drawing-containing top-level run
    return [r for r in p.findall('./w:r',NS) if r.findall('.//w:drawing',NS)]

def make_picture_para_from_runs(runs):
    p=ET.Element(q('p')); clear_pstyle(p); set_ppr(p,'center','0','240',False,True,True)
    for r in runs:
        nr=ET.fromstring(ET.tostring(r))
        # remove all text/tab/br from image run, keep drawing only
        for child in list(nr):
            if child.tag not in (q('rPr'), q('drawing')) and not child.findall('.//w:drawing',NS):
                nr.remove(child)
        set_run(nr,'宋体','Times New Roman',24,False,'000000')
        p.append(nr)
    return p

def remove_generated(body):
    removed=0
    for node in list(body):
        if node.tag==q('p'):
            t=text(node)
            if re.match(r'^图\d+\.\d+\s+图题$',t) or re.match(r'^表\d+\.\d+\s+表题$',t) or t=='注：资料来源于作者整理':
                body.remove(node); removed+=1
    return removed

def rebuild_body_picture_paragraphs(body):
    # Conservative mode: do not rebuild picture paragraphs unless text and image are mixed.
    # Rebuilding runs too aggressively can make later images disappear/turn blank after edits in Word.
    children=list(body); chapter=None; rebuilt=0; anchors=0; i=0
    while i < len(children):
        node=children[i]
        if node.tag==q('p') and is_h1(node):
            ch=h1_chapter(node)
            if ch is not None: chapter=ch
        if chapter is not None and node.tag==q('p') and has_drawing(node):
            anchors += convert_anchor_to_inline_in_para(node)
            txt=text(node)
            runs=drawing_runs_deep(node)
            if not runs:
                i+=1; continue
            # If this is already a pure picture paragraph, keep original XML intact.
            if not txt:
                i+=1; continue
            insert=[]
            insert.append(make_body_para(txt))
            insert.append(make_blank())
            insert.append(make_picture_para_from_runs(runs))
            idx=list(body).index(node)
            body.remove(node)
            for off,new in enumerate(insert): body.insert(idx+off,new)
            rebuilt+=1
            children=list(body); i=idx+len(insert); continue
        i+=1
    return rebuilt,anchors

FIG_CHAR = chr(0x56fe)
FIG_PLACEHOLDER_SUFFIX = chr(0x56fe) + chr(0x9898)

def is_figure_caption_text(t):
    return bool(re.match(r'^' + FIG_CHAR + r'\d+\.\d+\s*\S+', t or ''))

def is_generated_figure_placeholder(t):
    return bool(t and re.match(r'^' + FIG_CHAR + r'\d+\.\d+\s*' + FIG_PLACEHOLDER_SUFFIX + r'$', t))

def find_existing_figure_caption_around(body, pic_idx, scan_limit=8):
    children=list(body)
    for direction in (-1, 1):
        seen=0; j=pic_idx+direction
        while 0 <= j < len(children) and seen < scan_limit:
            n=children[j]
            if n.tag==q('tbl') or (n.tag==q('p') and is_h1(n)):
                break
            if n.tag==q('p'):
                tt=text(n)
                if is_figure_caption_text(tt) and not is_generated_figure_placeholder(tt):
                    return n
                if tt and not has_drawing(n) and not is_generated_figure_placeholder(tt):
                    break
            seen+=1; j+=direction
    return None

def format_existing_caption(p):
    set_ppr(p,'center','0','440',True,True,True)
    clear_pstyle(p)
    for r in p.findall('./w:r',NS): set_run(r,'宋体','Times New Roman',21,False,'000000')

def remove_node_if_present(body, node):
    try:
        body.remove(node); return True
    except ValueError:
        return False

def purge_extra_figure_captions_near(body, pic_node, keep_node=None, scan_limit=8):
    """Keep exactly one figure caption directly below the picture.
    Remove duplicate nearby real captions/placeholders, stopping at body text,
    another media/table block, or a heading.
    """
    removed=0; children=list(body)
    if pic_node not in children: return 0
    pic_idx=children.index(pic_node)
    for direction in (-1,1):
        seen=0; j=pic_idx+direction
        while 0 <= j < len(children) and seen < scan_limit:
            n=children[j]
            if n.tag==q('tbl') or (n.tag==q('p') and (is_h1(n) or has_drawing(n))):
                break
            if n.tag==q('p'):
                tt=text(n)
                if n is keep_node:
                    seen+=1; j+=direction; continue
                if is_figure_caption_text(tt) or is_generated_figure_placeholder(tt):
                    body.remove(n); removed+=1
                    children=list(body)
                    if pic_node not in children: return removed
                    pic_idx=children.index(pic_node); j=pic_idx+direction; seen=0
                    continue
                if tt:
                    break
            seen+=1; j+=direction
    return removed

def format_pictures_and_captions(body):
    children=list(body); chapter=None; seq={}; pics=0; caps=0; blanks=0; i=0
    while i < len(children):
        node=children[i]
        if node.tag==q('p') and is_h1(node):
            ch=h1_chapter(node)
            if ch is not None: chapter=ch
        if chapter is not None and node.tag==q('p') and has_drawing(node):
            pics+=1; seq[chapter]=seq.get(chapter,0)+1
            set_ppr(node,'center','0','240',False,True,True); clear_pstyle(node)
            for r in node.findall('./w:r',NS): set_run(r,'宋体','Times New Roman',24,False,'000000')
            idx=list(body).index(node)
            # Keep one blank before image if not already present.
            if idx==0 or not is_blank(list(body)[idx-1]):
                body.insert(idx,make_blank()); blanks+=1; idx+=1; node=list(body)[idx]
            else:
                set_ppr(list(body)[idx-1],'left','0','440',True)
            existing_cap=find_existing_figure_caption_around(body,idx)
            if existing_cap is not None:
                format_existing_caption(existing_cap)
                # 图题必须紧贴图片下方：图片段落之后立即是图题。
                remove_node_if_present(body, existing_cap)
                idx=list(body).index(node)
                body.insert(idx+1, existing_cap)
                cap_idx=idx+1
            else:
                cap_txt=f'{FIG_CHAR}{chapter}.{seq[chapter]} {FIG_PLACEHOLDER_SUFFIX}'
                body.insert(idx+1,make_caption(cap_txt)); caps+=1
                cap_idx=idx+1
            purge_extra_figure_captions_near(body, node, list(body)[cap_idx])
            cap_idx=list(body).index(list(body)[list(body).index(node)+1])
            # Keep one blank after the caption, not between image and caption.
            children=list(body); after=children[cap_idx+1] if cap_idx+1<len(children) else None
            if not is_blank(after):
                body.insert(cap_idx+1,make_blank()); blanks+=1
            else:
                set_ppr(after,'left','0','440',True)
            children=list(body); i=cap_idx+2; continue
        i+=1
    return pics,caps,blanks

# Reuse table logic by importing old script is messy; implement minimal here.
def clear_borders(borders):
    if borders is not None:
        for c in list(borders): borders.remove(c)
def set_border(parent,edge,val='single',sz='8'):
    e=parent.find(q(edge))
    if e is None: e=ET.SubElement(parent,q(edge))
    e.set(q('val'),val); e.set(q('sz'),sz); e.set(q('space'),'0'); e.set(q('color'),'000000')
def nil_border(parent,edge):
    e=parent.find(q(edge))
    if e is None: e=ET.SubElement(parent,q(edge))
    e.set(q('val'),'nil')

def none_border(parent,edge):
    e=parent.find(q(edge))
    if e is None: e=ET.SubElement(parent,q(edge))
    e.attrib.clear(); e.set(q('val'),'none'); e.set(q('color'),'auto'); e.set(q('sz'),'0'); e.set(q('space'),'0')
def format_table(tbl):
    tblPr=ensure(tbl,'tblPr',True)
    for tag in ('tblStyle','tblLook'):
        old=tblPr.find(q(tag))
        if old is not None:
            tblPr.remove(old)
    tb=tblPr.find(q('tblBorders'))
    if tb is None: tb=ET.SubElement(tblPr,q('tblBorders'))
    clear_borders(tb)
    for e in ('top','left','bottom','right','insideH','insideV'): nil_border(tb,e)
    rows=tbl.findall('./w:tr',NS)
    for ri,row in enumerate(rows):
        trPr=ensure(row,'trPr',True)
        if trPr.find(q('cantSplit')) is None: ET.SubElement(trPr,q('cantSplit'))
        for cell in row.findall('./w:tc',NS):
            tcPr=ensure(cell,'tcPr',True); cb=tcPr.find(q('tcBorders'))
            if cb is None: cb=ET.SubElement(tcPr,q('tcBorders'))
            clear_borders(cb)
            for e in ('left','right','insideH','insideV'): nil_border(cb,e)
            # Standard t0: all visible lines are cell borders.
            nil_border(cb,'top')
            nil_border(cb,'bottom')
            if ri==0:
                set_border(cb,'top','single','12')
                set_border(cb,'bottom','single','6')
            if ri==len(rows)-1:
                set_border(cb,'bottom','single','12')
            for p in cell.findall('.//w:p',NS):
                clear_pstyle(p)
                ppr=ensure(p,'pPr',True)
                # Verifier/school requirement: table cell text centered.
                jc=ppr.find(q('jc'))
                if jc is None:
                    jc=ET.SubElement(ppr,q('jc'))
                jc.set(q('val'),'center')
                ind=ensure(ppr,'ind')
                for a in ('firstLine','left','hanging','right','firstLineChars','leftChars','hangingChars','rightChars'):
                    ind.attrib.pop(q(a),None)
                sp=ensure(ppr,'spacing')
                sp.attrib.clear()
                sp.set(q('before'),'0'); sp.set(q('after'),'0'); sp.set(q('lineRule'),'auto')
                for r in p.findall('./w:r',NS): set_run(r,'宋体','Times New Roman',21,False,'000000')

TABLE_CHAR = chr(0x8868)
NOTE_CHAR = chr(0x6ce8)
TABLE_PLACEHOLDER_SUFFIX = chr(0x8868) + chr(0x9898)
NOTE_PREFIX = chr(0x6ce8) + chr(0xff1a)
NOTE_TEXT = ''.join(chr(x) for x in [0x6ce8,0xff1a,0x8d44,0x6599,0x6765,0x6e90,0x4e8e,0x4f5c,0x8005,0x6574,0x7406])

def zh(s):
    return s

def is_table_caption_text(t):
    return bool(re.match(r'^' + zh(TABLE_CHAR) + r'\d+\.\d+\s*\S+', t or ''))

def is_generated_table_placeholder(t):
    return bool(t and re.match(r'^' + zh(TABLE_CHAR) + r'\d+\.\d+\s*' + zh(TABLE_PLACEHOLDER_SUFFIX) + r'$', t))

def is_note_text(t):
    return bool(t and t.startswith(zh(NOTE_CHAR)))

def find_existing_table_caption_around(body, tbl_idx, scan_limit=8):
    """Find a nearby existing real table caption before OR after the table.
    Real captions may appear below the table in source docs. Preserve them and do
    not insert generated placeholders like 'Figure/Table placeholder'.
    """
    children=list(body)
    for direction in (-1, 1):
        seen=0; j=tbl_idx+direction
        while 0 <= j < len(children) and seen < scan_limit:
            n=children[j]
            if n.tag==q('tbl') or (n.tag==q('p') and is_h1(n)):
                break
            if n.tag==q('p'):
                tt=text(n)
                if is_table_caption_text(tt) and not is_generated_table_placeholder(tt):
                    return n
                if tt and not is_note_text(tt) and not has_drawing(n) and not is_generated_table_placeholder(tt):
                    break
            seen+=1; j+=direction
    return None

def remove_node_if_present(body, node):
    try:
        body.remove(node); return True
    except ValueError:
        return False

def normalize_table_caption(p):
    set_ppr(p,'center','0','440',True,True,True)
    clear_pstyle(p)
    for r in p.findall('./w:r',NS): set_run(r,'宋体','Times New Roman',21,False,'000000')


def keep_table_with_next_paragraphs(tbl):
    # Hard guard against table pagination: Word has no single "keep whole table"
    # flag, so combine row cantSplit with paragraph keepLines/keepNext.
    rows=tbl.findall('./w:tr',NS)
    for ri,row in enumerate(rows):
        trPr=ensure(row,'trPr',True)
        if trPr.find(q('cantSplit')) is None:
            ET.SubElement(trPr,q('cantSplit'))
        for p in row.findall('.//w:p',NS):
            ppr=ensure(p,'pPr',True)
            if ppr.find(q('keepLines')) is None:
                ET.SubElement(ppr,q('keepLines'))
            kn=ppr.find(q('keepNext'))
            if ri < len(rows)-1 and kn is None:
                ET.SubElement(ppr,q('keepNext'))
            elif ri == len(rows)-1 and kn is not None:
                ppr.remove(kn)
def format_tables(body):
    children=list(body); chapter=None; seq={}; tbls=0; caps=0; notes=0; blanks=0; i=0
    while i < len(children):
        node=children[i]
        if node.tag==q('p') and is_h1(node):
            ch=h1_chapter(node)
            if ch is not None: chapter=ch
        if chapter is not None and node.tag==q('tbl'):
            tbls+=1; seq[chapter]=seq.get(chapter,0)+1; format_table(node); keep_table_with_next_paragraphs(node)
            idx=list(body).index(node)
            existing_cap=find_existing_table_caption_around(body,idx)

            # 表格前后与正文空一行；表题必须在表格上方，且紧贴表格。
            if existing_cap is not None:
                normalize_table_caption(existing_cap)
                # Remove caption from old location (above or below) and reinsert directly before table.
                remove_node_if_present(body, existing_cap)
                idx=list(body).index(node)
                # Ensure one blank before caption unless already blank.
                if idx>0 and not is_blank(list(body)[idx-1]):
                    body.insert(idx,make_blank()); blanks+=1; idx+=1
                body.insert(idx, existing_cap); idx+=1
            else:
                if idx>0 and not is_blank(list(body)[idx-1]):
                    body.insert(idx,make_blank()); blanks+=1; idx+=1
                cap_txt=f'{zh(TABLE_CHAR)}{chapter}.{seq[chapter]} {zh(TABLE_PLACEHOLDER_SUFFIX)}'
                body.insert(idx,make_caption(cap_txt)); caps+=1; idx+=1

            # idx now points to table after caption insertion.
            idx=list(body).index(node)
            # School rule: keep the table block integral, with one blank line after the table.
            # Do not auto-insert fabricated source notes unless they already exist in source.
            children=list(body); after=children[idx+1] if idx+1<len(children) else None
            if after is not None and after.tag==q('p') and is_note_text(text(after)):
                set_ppr(after,'left','0','440',True,False,True)
                idx=idx+1
            children=list(body); after2=children[idx+1] if idx+1<len(children) else None
            if not is_blank(after2): body.insert(idx+1,make_blank()); blanks+=1
            children=list(body); i=idx+2; continue
        i+=1
    return tbls,caps,notes,blanks

def process(src,out):
    src=Path(src); out=Path(out); out.parent.mkdir(parents=True,exist_ok=True); tmp=Path(tempfile.mkdtemp(prefix='step18v3_'))
    try:
        with ZipFile(src) as z: z.extractall(tmp)
        docxml=tmp/'word'/'document.xml'; tree=ET.parse(str(docxml)); root=tree.getroot(); body=root.find(q('body'))
        removed=remove_generated(body)
        rebuilt,anchors=rebuild_body_picture_paragraphs(body)
        pics,caps,pic_blanks=format_pictures_and_captions(body)
        tbls,tbl_caps,notes,tbl_blanks=format_tables(body)
        tree.write(str(docxml),xml_declaration=True,encoding='utf-8',standalone='yes')
        repack=out.with_suffix('.tmp.docx')
        with ZipFile(repack,'w',ZIP_DEFLATED) as z:
            for f in tmp.rglob('*'):
                if f.is_file(): z.write(f,f.relative_to(tmp))
        shutil.move(str(repack),str(out))
        print('OUTPUT:',out)
        print('REMOVED_OLD_GENERATED:',removed)
        print('TEXT_IMAGE_PARAS_REBUILT:',rebuilt)
        print('ANCHOR_TO_INLINE_BODY:',anchors)
        print('BODY_PICTURES:',pics,'PIC_CAPTIONS:',caps,'PIC_BLANKS:',pic_blanks)
        print('BODY_TABLES:',tbls,'TABLE_CAPTIONS:',tbl_caps,'TABLE_NOTES:',notes,'TABLE_BLANKS:',tbl_blanks)
    finally: shutil.rmtree(tmp,ignore_errors=True)
if __name__=='__main__':
    if len(sys.argv)!=3: raise SystemExit('usage: python step18_xml_fix_media_tables_v3.py in.docx out.docx')
    process(sys.argv[1],sys.argv[2])



