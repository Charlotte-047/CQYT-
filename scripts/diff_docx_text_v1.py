from zipfile import ZipFile
from pathlib import Path
from lxml import etree as ET
import sys,re,json,difflib
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS={'w':W.strip('{}')}
def q(t): return f'{W}{t}'
def txt(p): return ''.join(t.text or '' for t in p.findall('.//w:t',NS)).strip()
def norm(t): return re.sub(r'\s+','',t or '')
def load(path):
    root=ET.fromstring(ZipFile(path).read('word/document.xml'))
    return [txt(p) for p in root.findall('.//w:body/w:p',NS) if txt(p)]
def main(a,b):
    A=load(a); B=load(b)
    nA=[norm(x) for x in A]; nB=[norm(x) for x in B]
    setB='\n'.join(nB)
    missing=[]
    for i,x in enumerate(A):
        nx=nA[i]
        if len(nx)>=20 and nx not in setB:
            missing.append((i,x[:160]))
    added=[]
    setA='\n'.join(nA)
    for i,x in enumerate(B):
        nx=nB[i]
        if len(nx)>=20 and nx not in setA:
            # likely generated TOC rows if ends with page number and starts as heading; keep but classify
            cls='toc_or_generated' if re.match(r'^(\d+(?:\.\d+)*|致谢|参考文献).+\d$', x) else 'new_or_modified'
            added.append((i,cls,x[:160]))
    print('SRC_NONEMPTY',len(A))
    print('OUT_NONEMPTY',len(B))
    print('MISSING_LONG_COUNT',len(missing))
    for m in missing[:30]: print('MISSING',m)
    print('ADDED_OR_CHANGED_LONG_COUNT',len(added))
    for x in added[:80]: print('ADDED',x)
    print('RESULT', '通过' if not missing else '未通过')
    return 0 if not missing else 1
if __name__=='__main__': raise SystemExit(main(Path(sys.argv[1]),Path(sys.argv[2])))
