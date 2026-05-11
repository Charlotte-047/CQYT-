from zipfile import ZipFile
from pathlib import Path
from lxml import etree as ET
import sys,re,json
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS={'w':W.strip('{}')}
def q(t): return f'{W}{t}'
def txt(p): return ''.join(t.text or '' for t in p.findall('.//w:t',NS)).strip()
def norm(t): return re.sub(r'\s+','',t or '')
def style(p):
    ppr=p.find(q('pPr')); ps=None if ppr is None else ppr.find(q('pStyle'))
    return '' if ps is None else (ps.get(q('val')) or '')
def load(path):
    root=ET.fromstring(ZipFile(path).read('word/document.xml'))
    return root.findall('.//w:body/w:p',NS)
def is_toc_title(t): return norm(t)==(chr(0x76ee)+chr(0x5f55))
def is_toc_entry(t): return bool(re.match(r'^(\d+(?:\.\d+)*\s+.+|'+chr(0x81f4)+chr(0x8c22)+'|'+chr(0x53c2)+chr(0x8003)+chr(0x6587)+chr(0x732e)+r')\s*\d+$',t))
def body_like(t): return len(t)>90 or bool(re.search(r'[。；;]$',t))
def hlevel(t):
    if body_like(t): return None
    if norm(t) in (chr(0x81f4)+chr(0x8c22), chr(0x53c2)+chr(0x8003)+chr(0x6587)+chr(0x732e)): return 1
    if re.match(r'^\d+\.\d+\.\d+\s*\S.{0,80}$',t): return 3
    if re.match(r'^\d+\.\d+\s*\S.{0,75}$',t): return 2
    if re.match(r'^\d+\s+\S.{0,60}$',t): return 1
    return None
def toc_entries(ps):
    blocks=[]; cur=None
    for i,p in enumerate(ps):
        t=txt(p)
        if is_toc_title(t):
            if cur: blocks.append(cur)
            cur={'title_idx':i,'entries':[]}; continue
        if cur is not None:
            if is_toc_entry(t) or style(p).startswith('TOC'):
                if t: cur['entries'].append(t)
                continue
            if hlevel(t)==1:
                blocks.append(cur); cur=None
            elif t=='':
                continue
    if cur: blocks.append(cur)
    return blocks
def main(src,out):
    sblocks=toc_entries(load(src)); oblocks=toc_entries(load(out))
    fails=[]; warns=[]
    if not oblocks: fails.append('output toc missing')
    if len(oblocks)>1: fails.append(f'output has multiple toc blocks: {len(oblocks)}')
    if sblocks and oblocks:
        se=sblocks[0]['entries']; oe=oblocks[0]['entries']
        if len(oe) < max(3, len(se)//2): fails.append(f'output toc too few entries: src={len(se)} out={len(oe)}')
        # Preserve real page-numbered TOC entries where source had them.
        smap={re.sub(r'\d+$','',norm(x)): re.search(r'(\d+)$',norm(x)).group(1) for x in se if re.search(r'\d+$',norm(x))}
        omap={re.sub(r'\d+$','',norm(x)): re.search(r'(\d+)$',norm(x)).group(1) for x in oe if re.search(r'\d+$',norm(x))}
        changed=[]
        for k,v in smap.items():
            if k in omap and omap[k]!=v:
                changed.append((k,v,omap[k]))
        if changed: fails.append(f'toc page numbers changed examples={changed[:12]} count={len(changed)}')
    print('TOC_PRESERVED_SRC_BLOCKS',len(sblocks),'OUT_BLOCKS',len(oblocks))
    print('TOC_PRESERVED_SRC_ENTRIES',len(sblocks[0]['entries']) if sblocks else 0,'OUT_ENTRIES',len(oblocks[0]['entries']) if oblocks else 0)
    print('WARNING_COUNT',len(warns))
    for w in warns: print('WARN',w)
    print('FAILURE_COUNT',len(fails))
    for f in fails: print('FAIL',f)
    print('RESULT','通过' if not fails else '未通过')
    return 0 if not fails else 1
if __name__=='__main__': raise SystemExit(main(Path(sys.argv[1]),Path(sys.argv[2])))
