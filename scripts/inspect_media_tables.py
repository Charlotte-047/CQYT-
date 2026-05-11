from zipfile import ZipFile
from pathlib import Path
from lxml import etree as ET
import sys
p=Path(sys.argv[1])
NS={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main','wp':'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
with ZipFile(p) as z:
    names=z.namelist()
    root=ET.fromstring(z.read('word/document.xml'))
    rels=z.read('word/_rels/document.xml.rels').decode('utf-8',errors='ignore')
print('FILE',p)
print('headers', [n for n in names if n.startswith('word/header')])
print('footers', [n for n in names if n.startswith('word/footer')])
print('tables', len(root.findall('.//w:tbl',NS)))
print('drawings', len(root.findall('.//w:drawing',NS)))
print('inline drawings', len(root.findall('.//wp:inline',NS)))
print('anchor drawings', len(root.findall('.//wp:anchor',NS)))
for i,pgh in enumerate(root.findall('.//w:p',NS),1):
    txt=''.join(t.text or '' for t in pgh.findall('.//w:t',NS)).strip()
    if pgh.findall('.//w:drawing',NS) or txt.startswith(('图','表')):
        print(i, repr(txt), 'draw', bool(pgh.findall('.//w:drawing',NS)), 'tabs', len(pgh.findall('.//w:tab',NS)))
