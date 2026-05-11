from zipfile import ZipFile, ZIP_DEFLATED
from pathlib import Path
from lxml import etree as ET
import tempfile, shutil, sys, re, copy

W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS={'w':W.strip('{}')}
def q(t): return f'{W}{t}'
def text(p): return ''.join(t.text or '' for t in p.findall('.//w:t',NS)).strip()
def norm(t): return re.sub(r'\s+','',t or '')
def style(p):
    ppr=p.find(q('pPr')); ps=None if ppr is None else ppr.find(q('pStyle'))
    return '' if ps is None else (ps.get(q('val')) or '')
def is_toc_title(t): return norm(t)==(chr(0x76ee)+chr(0x5f55))
def body_like(t): return len(t)>90 or bool(re.search(r'[。；;]$',t))
def hlevel(t):
    if body_like(t): return None
    if norm(t) in (chr(0x81f4)+chr(0x8c22), chr(0x53c2)+chr(0x8003)+chr(0x6587)+chr(0x732e)): return 1
    if re.match(r'^\d+\.\d+\.\d+\s*\S.{0,80}$',t): return 3
    if re.match(r'^\d+\.\d+\s*\S.{0,75}$',t): return 2
    if re.match(r'^\d+\s+\S.{0,60}$',t): return 1
    return None

def is_toc_like(p):
    t=text(p); st=style(p)
    if st.startswith('TOC') or st in ('17','26','31','193'):
        # source document uses numeric styles for TOC title/levels
        return True
    if bool(re.match(r'^(\d+(?:\.\d+)*\s+.+|'+chr(0x81f4)+chr(0x8c22)+'|'+chr(0x53c2)+chr(0x8003)+chr(0x6587)+chr(0x732e)+r')\s*\d+$',t)):
        return True
    # field TOC rows may include instrText even when text not matching due display truncation
    instr=''.join(x.text or '' for x in p.findall('.//w:instrText',NS))
    return 'TOC' in instr or 'PAGEREF' in instr or 'HYPERLINK' in instr

def extract_toc(nodes):
    for i,n in enumerate(nodes):
        if n.tag==q('p') and is_toc_title(text(n)):
            j=i+1
            while j<len(nodes):
                m=nodes[j]
                if m.tag==q('p'):
                    mt=text(m)
                    if hlevel(mt)==1 and not is_toc_like(m):
                        break
                    # include TOC rows and separating blanks; stop only at real body H1
                    j+=1; continue
                j+=1
            return i,j,[copy.deepcopy(x) for x in nodes[i:j]]
    return None

def remove_toc(nodes):
    ext=extract_toc(nodes)
    if not ext: return nodes,0,None
    i,j,block=ext
    return nodes[:i]+nodes[j:], j-i, i

def set_pagebreak_before(p,on=True):
    ppr=p.find(q('pPr'))
    if ppr is None:
        ppr=ET.Element(q('pPr')); p.insert(0,ppr)
    for n in list(ppr.findall(q('pageBreakBefore'))): ppr.remove(n)
    if on: ppr.insert(0,ET.Element(q('pageBreakBefore')))

def main(src,out):
    src=Path(src); out=Path(out); tmp=Path(tempfile.mkdtemp())
    try:
        src_root=ET.fromstring(ZipFile(src).read('word/document.xml'))
        src_nodes=[n for n in list(src_root.find(q('body'))) if n.tag!=q('sectPr')]
        ext=extract_toc(src_nodes)
        if not ext: raise RuntimeError('source TOC block not found')
        _,_,src_toc=ext
        with ZipFile(out) as z: z.extractall(tmp)
        docxml=tmp/'word'/'document.xml'
        tree=ET.parse(str(docxml)); root=tree.getroot(); body=root.find(q('body'))
        sect=[n for n in list(body) if n.tag==q('sectPr')]
        nodes=[n for n in list(body) if n.tag!=q('sectPr')]
        nodes,removed,old_idx=remove_toc(nodes)
        # Insert source TOC before first real body H1, or old TOC position if no H1 found.
        insert_idx=next((i for i,n in enumerate(nodes) if n.tag==q('p') and hlevel(text(n))==1), old_idx if old_idx is not None else len(nodes))
        nodes=nodes[:insert_idx]+[copy.deepcopy(x) for x in src_toc]+nodes[insert_idx:]
        # Keep body first H1 page break; do not touch TOC internals.
        if insert_idx+len(src_toc)<len(nodes) and nodes[insert_idx+len(src_toc)].tag==q('p'):
            set_pagebreak_before(nodes[insert_idx+len(src_toc)], True)
        body[:] = nodes+sect
        tree.write(str(docxml),xml_declaration=True,encoding='utf-8',standalone='yes')
        bak=out.with_suffix(out.suffix+'.bak_restore_source_toc')
        shutil.copyfile(out,bak)
        with ZipFile(out,'w',ZIP_DEFLATED) as z:
            for f in tmp.rglob('*'):
                if f.is_file(): z.write(f,f.relative_to(tmp).as_posix())
        print('RESTORE_SOURCE_TOC_SRC_BLOCK',len(src_toc))
        print('RESTORE_SOURCE_TOC_REMOVED_OUT_BLOCK',removed)
        print('RESTORE_SOURCE_TOC_BACKUP',bak)
    finally:
        shutil.rmtree(tmp,ignore_errors=True)
if __name__=='__main__': main(sys.argv[1], sys.argv[2])
