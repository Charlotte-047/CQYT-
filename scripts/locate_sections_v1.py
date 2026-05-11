from zipfile import ZipFile
from lxml import etree as ET
from pathlib import Path
import sys,re
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS={'w':W.strip('{}')}
def q(t): return f'{W}{t}'
def text(p): return ''.join(t.text or '' for t in p.findall('.//w:t',NS)).strip()
def style(p):
    ppr=p.find(q('pPr')); ps=None if ppr is None else ppr.find(q('pStyle'))
    return '' if ps is None else (ps.get(q('val')) or '')
def norm(t): return re.sub(r'\s+','',t or '')
def hit(t):
    n=norm(t)
    return any(k in n for k in ['原创性','摘要','ABSTRACT','目录','绪论','参考文献','致谢']) or bool(re.match(r'^\d+(?:\.\d+)*\s*\S+',t or ''))
with ZipFile(sys.argv[1]) as z: root=ET.fromstring(z.read('word/document.xml'))
paras=root.find(q('body')).findall('./w:p',NS)
for i,p in enumerate(paras):
    t=text(p)
    if hit(t): print(i,'STYLE',style(p),'NORM',norm(t),'TEXT',repr(t[:120]))
