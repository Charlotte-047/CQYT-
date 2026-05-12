from zipfile import ZipFile, ZIP_DEFLATED
from pathlib import Path
from lxml import etree as ET
import sys, tempfile, shutil, re

W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS={'w':W.strip('{}')}

def q(t): return f'{W}{t}'
def text(p): return ''.join(t.text or '' for t in p.findall('.//w:t',NS)).strip()
def style(p):
    ppr=p.find(q('pPr')); ps=None if ppr is None else ppr.find(q('pStyle'))
    return '' if ps is None else ps.get(q('val')) or ''
def ensure(parent, tag, first=False):
    n=parent.find(q(tag))
    if n is None:
        n=ET.Element(q(tag))
        if first: parent.insert(0,n)
        else: parent.append(n)
    return n

def clear_runs(p):
    for r in list(p.findall('./w:r',NS)): p.remove(r)

def set_run_fonts(r, east='宋体', ascii='Times New Roman', size=21, bold=False, color='000000'):
    rpr=ensure(r,'rPr',True)
    for c in list(rpr):
        if c.tag in (q('rFonts'),q('sz'),q('szCs'),q('b'),q('bCs'),q('color'),q('u')): rpr.remove(c)
    rf=ET.SubElement(rpr,q('rFonts'))
    rf.set(q('eastAsia'),east); rf.set(q('ascii'),ascii); rf.set(q('hAnsi'),ascii); rf.set(q('cs'),ascii)
    b=ET.SubElement(rpr,q('b')); b.set(q('val'),'1' if bold else '0')
    bcs=ET.SubElement(rpr,q('bCs')); bcs.set(q('val'),'1' if bold else '0')
    sz=ET.SubElement(rpr,q('sz')); sz.set(q('val'),str(size))
    szcs=ET.SubElement(rpr,q('szCs')); szcs.set(q('val'),str(size))
    c=ET.SubElement(rpr,q('color')); c.set(q('val'),color)

def is_cjk(ch):
    return '\u4e00' <= ch <= '\u9fff'

def add_text_run(p, s, cjk):
    if not s: return
    r=ET.SubElement(p,q('r'))
    # Non-Chinese text, including digits, English and punctuation, must be Times New Roman.
    if cjk:
        set_run_fonts(r,'宋体','Times New Roman',21,False,'000000')
    else:
        set_run_fonts(r,'Times New Roman','Times New Roman',21,False,'000000')
    t=ET.SubElement(r,q('t')); t.text=s
    if s.startswith(' ') or s.endswith(' '):
        t.set('{http://www.w3.org/XML/1998/namespace}space','preserve')

def replace_mixed_font_text(p, s):
    clear_runs(p)
    buf=''; cur=None
    for ch in s:
        cjk=is_cjk(ch)
        if cur is None:
            buf=ch; cur=cjk
        elif cjk==cur:
            buf+=ch
        else:
            add_text_run(p,buf,cur); buf=ch; cur=cjk
    add_text_run(p,buf,cur if cur is not None else False)

def set_ref_title(p):
    ppr=ensure(p,'pPr',True)
    ps=ensure(ppr,'pStyle',True); ps.set(q('val'),'2')
    jc=ensure(ppr,'jc'); jc.set(q('val'),'center')
    sp=ensure(ppr,'spacing'); sp.set(q('before'),'0'); sp.set(q('after'),'0'); sp.set(q('line'),'440'); sp.set(q('lineRule'),'exact')
    ind=ensure(ppr,'ind')
    for a in ('firstLine','left','hanging','right','firstLineChars','leftChars','hangingChars','rightChars'):
        ind.attrib.pop(q(a),None)
    clear_runs(p)
    r=ET.SubElement(p,q('r')); set_run_fonts(r,'黑体','Times New Roman',32,False,'000000')
    ET.SubElement(r,q('t')).text='参考文献'

def set_ref_ppr(p):
    ppr=ensure(p,'pPr',True)
    ps=ppr.find(q('pStyle'))
    if ps is not None: ppr.remove(ps)
    jc=ensure(ppr,'jc'); jc.set(q('val'),'both')
    sp=ensure(ppr,'spacing'); sp.set(q('before'),'0'); sp.set(q('after'),'0'); sp.set(q('line'),'440'); sp.set(q('lineRule'),'exact')
    ind=ensure(ppr,'ind')
    # 规范：序号左顶格。这里清掉 left/hanging/firstLine，避免悬挂缩进。
    for a in ('firstLine','left','hanging','right','firstLineChars','leftChars','hangingChars','rightChars'):
        ind.attrib.pop(q(a),None)
    for tag in ('keepNext','keepLines'):
        old=ppr.find(q(tag))
        if old is not None: ppr.remove(old)

def normalize_punct(s):
    """Preserve reference text content.

    Earlier versions converted Chinese punctuation/quotes to English punctuation,
    which changed bibliography text and triggered diff failures. Formatting scripts
    must not rewrite user content; only normalize reference numbering whitespace.
    """
    s=s.replace('\t',' ')
    s=re.sub(r'\s+',' ',s).strip()
    m=re.match(r'^\[?0*(\d+)\]?\s*[\.、．]?\s*(.*)$',s)
    if m:
        s=f'[{int(m.group(1))}] {m.group(2).strip()}'
    return s

def is_ref_line(t):
    return bool(re.match(r'^\s*\[?\d+\]?\s*[\.、．\t ]+',t))

def process(src,out):
    src=Path(src); out=Path(out); out.parent.mkdir(parents=True,exist_ok=True); tmp=Path(tempfile.mkdtemp(prefix='step19refv2_'))
    try:
        with ZipFile(src) as z: z.extractall(tmp)
        docxml=tmp/'word'/'document.xml'; tree=ET.parse(str(docxml)); root=tree.getroot(); body=root.find(q('body'))
        in_refs=False; refs=0; normalized=0; title_fixed=0; blanks_removed=0
        for p in list(body.findall('./w:p',NS)):
            t=text(p)
            if t=='参考文献':
                in_refs=True; title_fixed+=1; set_ref_title(p); continue
            if not in_refs: continue
            if t and style(p) in ('2','3','Heading1','Heading2') and t!='参考文献':
                in_refs=False; continue
            if not t:
                idx=list(body).index(p); prev=list(body)[idx-1] if idx>0 else None; nxt=list(body)[idx+1] if idx+1<len(list(body)) else None
                if prev is not None and text(prev)=='参考文献':
                    ppr=ensure(p,'pPr',True); jc=ensure(ppr,'jc'); jc.set(q('val'),'left')
                    sp=ensure(ppr,'spacing'); sp.set(q('before'),'0'); sp.set(q('after'),'0'); sp.set(q('line'),'440'); sp.set(q('lineRule'),'exact')
                    ind=ensure(ppr,'ind')
                    for a in ('firstLine','left','hanging','right','firstLineChars','leftChars','hangingChars','rightChars'): ind.attrib.pop(q(a),None)
                    continue
                if nxt is None or not text(nxt):
                    body.remove(p); blanks_removed+=1
                continue
            if is_ref_line(t) or refs>0:
                refs+=1
                new=normalize_punct(t)
                # If numbering was absent/broken, enforce sequence.
                new=re.sub(r'^\[\d+\]', f'[{refs}]', new)
                if new != t: normalized+=1
                set_ref_ppr(p)
                replace_mixed_font_text(p,new)
        tree.write(str(docxml),xml_declaration=True,encoding='utf-8',standalone='yes')
        repack=out.with_suffix('.tmp.docx')
        with ZipFile(repack,'w',ZIP_DEFLATED) as z:
            for f in tmp.rglob('*'):
                if f.is_file(): z.write(f,f.relative_to(tmp))
        shutil.move(str(repack),str(out))
        print('OUTPUT:',out)
        print('REFERENCE_TITLE_FIXED:',title_fixed)
        print('REFERENCES_FORMATTED:',refs)
        print('REFERENCES_NORMALIZED:',normalized)
        print('BLANKS_REMOVED:',blanks_removed)
    finally: shutil.rmtree(tmp,ignore_errors=True)
if __name__=='__main__':
    if len(sys.argv)!=3: raise SystemExit('usage: python step19_xml_fix_references_v2.py in.docx out.docx')
    process(sys.argv[1],sys.argv[2])
