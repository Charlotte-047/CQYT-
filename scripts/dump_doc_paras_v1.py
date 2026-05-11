from zipfile import ZipFile
from lxml import etree as ET
from pathlib import Path
import sys
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS={'w':W.strip('{}')}
def q(t): return f'{W}{t}'
def text(p): return ''.join(t.text or '' for t in p.findall('.//w:t',NS)).strip()
def style(p):
 ppr=p.find(q('pPr')); ps=None if ppr is None else ppr.find(q('pStyle'))
 return '' if ps is None else (ps.get(q('val')) or '')
with ZipFile(sys.argv[1]) as z: root=ET.fromstring(z.read('word/document.xml'))
body=root.find(q('body'))
for i,p in enumerate(body.findall('./w:p',NS)):
 t=text(p)
 if t or i<80:
  print(i, 'STYLE', style(p), 'TEXT', repr(t[:120]))
