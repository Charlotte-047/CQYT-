from zipfile import ZipFile
from pathlib import Path
from lxml import etree as ET
import sys,re,json

W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
WP='{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
NS={'w':W.strip('{}'),'wp':WP.strip('{}')}
def q(t): return f'{W}{t}'
def texts(root): return [t.text or '' for t in root.findall('.//w:t',NS)]
def ptext(p): return ''.join(t.text or '' for t in p.findall('.//w:t',NS)).strip()
def parse(z,name): return ET.fromstring(z.read(name))
def attrdict(n): return {} if n is None else {k.split('}')[-1]:v for k,v in n.attrib.items()}
def main(path):
    path=Path(path); fails=[]; warns=[]; metrics={}
    with ZipFile(path) as z:
        names=z.namelist()
        xml_names=[n for n in names if n.startswith('word/') and n.endswith('.xml')]
        roots={n:parse(z,n) for n in xml_names if n in names}
    bad_text=[]
    suspicious_chars=[chr(0xfffd), chr(0x25a1)]
    for name,root in roots.items():
        for tx in texts(root):
            if any(c in tx for c in suspicious_chars):
                bad_text.append((name,tx[:120]))
    # question marks alone are legal; only flag repeated question patterns near Chinese-like contexts.
    if bad_text:
        fails.append(f'suspicious replacement/garbled text runs: {len(bad_text)} examples={bad_text[:20]}')
    doc=roots['word/document.xml']; body=doc.find(q('body'))
    paras=body.findall('./w:p',NS)
    sects=doc.findall('.//w:sectPr',NS)
    metrics['sections']=len(sects)
    metrics['paragraphs']=len(paras)
    # Section/page numbering summary.
    sect_info=[]
    for i,s in enumerate(sects,1):
        pg=s.find(q('pgNumType'))
        sect_info.append({'idx':i,'pgNumType':attrdict(pg),'headers':len(s.findall(q('headerReference'))),'footers':len(s.findall(q('footerReference')))})
    metrics['section_info']=sect_info
    if len(sects)<3: warns.append(f'few sections detected: {len(sects)}')
    # Page breaks before H1.
    def body_like(t): return len(t)>70 or bool(re.search(r'[。；;]$',t))
    def hlevel(t):
        if body_like(t): return None
        if t in ('致谢','参考文献'): return 1
        if re.match(r'^\d+\.\d+\.\d+\s*\S.{0,70}$',t): return 3
        if re.match(r'^\d+\.\d+\s*\S.{0,65}$',t): return 2
        if re.match(r'^\d+\s+\S.{0,55}$',t): return 1
        return None
    def has_pbb(p):
        ppr=p.find(q('pPr')); return ppr is not None and ppr.find(q('pageBreakBefore')) is not None
    def has_pagebr(p): return bool(p.findall('.//w:br[@w:type="page"]',NS))
    h1=[]
    in_toc=False
    for i,p in enumerate(paras):
        t=ptext(p); nt=re.sub(r'\s+','',t)
        if nt=='目录': in_toc=True; continue
        if in_toc:
            if re.match(r'^(\d+(?:\.\d+)*\s+.+|致谢|参考文献)\s*\d+$',t): continue
            if hlevel(t)==1: in_toc=False
            else: continue
        if hlevel(t)==1:
            h1.append((i,t,has_pbb(p),has_pagebr(p)))
    metrics['h1_pagebreaks']=h1
    missing_h1=[x for x in h1 if not x[2]]
    if missing_h1: fails.append(f'H1 missing pageBreakBefore: {missing_h1[:10]}')
    # Drawings: anchor vs inline.
    anchors=doc.findall('.//wp:anchor',NS); inlines=doc.findall('.//wp:inline',NS)
    metrics['drawing_anchors']=len(anchors); metrics['drawing_inlines']=len(inlines)
    # One anchor in header/cover may be acceptable; body floating anchors are high risk.
    body_anchor_paras=[]
    for i,p in enumerate(paras):
        if p.findall('.//wp:anchor',NS): body_anchor_paras.append((i,ptext(p)[:80]))
    metrics['body_anchor_paragraphs']=body_anchor_paras
    if body_anchor_paras: warns.append(f'body floating anchors found: {body_anchor_paras[:10]}')
    # Header/footer files exist and not obviously garbled.
    headers=[n for n in roots if n.startswith('word/header')]
    footers=[n for n in roots if n.startswith('word/footer')]
    metrics['headers']=headers; metrics['footers']=footers
    if not footers: warns.append('no footer xml found')
    # Required front matter markers.
    all_text='\n'.join(ptext(p) for p in paras)
    for marker in ['原创性声明','摘  要','ABSTRACT','目    录','1 绪论','参考文献']:
        if marker not in all_text:
            warns.append(f'front/body marker not exact-found: {marker}')
    print('DOCX_INTEGRITY_METRICS',json.dumps(metrics,ensure_ascii=False))
    print('WARNING_COUNT',len(warns))
    for w in warns: print('WARN',w)
    print('FAILURE_COUNT',len(fails))
    for f in fails: print('FAIL',f)
    print('RESULT','通过' if not fails else '未通过')
    return 0 if not fails else 1
if __name__=='__main__': raise SystemExit(main(sys.argv[1]))
