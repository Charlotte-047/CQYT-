from zipfile import ZipFile
from lxml import etree as ET
import sys,re
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS={'w':W.strip('{}')}
def q(t): return f'{W}{t}'
def text(p): return ''.join(t.text or '' for t in p.findall('.//w:t',NS)).strip()
def raw(p): return ''.join(t.text or '' for t in p.findall('.//w:t',NS))
def has_drawing(p): return bool(p.findall('.//w:drawing',NS))
def norm(t): return ''.join((t or '').split())
def is_toc_title(t): return norm(t)==(chr(0x76ee)+chr(0x5f55))
def is_toc_entry_text(t):
 import re
 return bool(re.match(r'^(\d+(?:\.\d+)*\s+.+|'+chr(0x81f4)+chr(0x8c22)+'|'+chr(0x53c2)+chr(0x8003)+chr(0x6587)+chr(0x732e)+r')\s*\d+$',t or ''))
def h1_like(t):
 import re
 tt=t or ''
 if len(tt)>90 or re.search(r'['+chr(0x3002)+chr(0xff1b)+r';]$',tt): return False
 return bool(re.match(r'^\d+\s+\S.{0,60}$',tt)) or norm(tt) in (chr(0x81f4)+chr(0x8c22), chr(0x53c2)+chr(0x8003)+chr(0x6587)+chr(0x732e))
def blank_expected_font(paras, idx):
 # English abstract/front-matter blanks use Times New Roman; Chinese/body blanks use 宋体.
 state='front'
 for j in range(0, idx+1):
  tt=text(paras[j])
  if tt=='ABSTRACT': state='abstract_en'
  elif norm(tt)==(chr(0x76ee)+chr(0x5f55)): state='toc'
  elif re.match(r'^1\s+\S+', tt or ''): state='body'
 return 'Times New Roman' if state=='abstract_en' else '宋体'

def fonts(p):
 out=[]
 ppr=p.find(q('pPr'))
 if ppr is not None:
  rpr=ppr.find(q('rPr'))
  if rpr is not None:
   rf=rpr.find(q('rFonts')); sz=rpr.find(q('sz'))
   out.append(('pmark', None if rf is None else {k.split('}')[-1]:v for k,v in rf.attrib.items()}, None if sz is None else sz.get(q('val'))))
 for r in p.findall('./w:r',NS):
  rpr=r.find(q('rPr'))
  if rpr is not None:
   rf=rpr.find(q('rFonts')); sz=rpr.find(q('sz'))
   out.append(('run', None if rf is None else {k.split('}')[-1]:v for k,v in rf.attrib.items()}, None if sz is None else sz.get(q('val'))))
 return out
def one_font(fdict,name):
 if not fdict: return False
 if name=='Times New Roman':
  return all(fdict.get(k)=='Times New Roman' for k in ('eastAsia','ascii','hAnsi','cs'))
 if name=='宋体':
  all_song=all(fdict.get(k)=='宋体' for k in ('eastAsia','ascii','hAnsi','cs'))
  mixed=fdict.get('eastAsia')=='宋体' and fdict.get('ascii')=='Times New Roman' and fdict.get('hAnsi')=='Times New Roman'
  return all_song or mixed
 return all(fdict.get(k)==name for k in ('eastAsia','ascii','hAnsi','cs'))
p=sys.argv[1]; fails=[]
with ZipFile(p) as z:
 root=ET.fromstring(z.read('word/document.xml'))
 ps=root.findall('.//w:p',NS)
 abstract_idx=None
 for i,pa in enumerate(ps):
  if text(pa).upper()=='ABSTRACT': abstract_idx=i
 for i,pa in enumerate(ps):
  if pa.getparent().tag==q('body') and ((text(pa)=='' and raw(pa).strip()=='') or raw(pa)==' ') and not has_drawing(pa):
   expected=blank_expected_font(ps, i)
   fs=fonts(pa)
   valid_pmark=any(x[0].lower()=='pmark' and (one_font(x[1],expected) or (expected==chr(0x5b8b)+chr(0x4f53) and (x[1].get('eastAsia')==chr(0x5b8b)+chr(0x4f53) or x[1].get('cs')==chr(0x5b8b)+chr(0x4f53)))) and (x[2] in ('24', None)) for x in fs)
   run_fs=[x for x in fs if x[0].lower()=='run' and x[2] is not None] or ([] if valid_pmark else fs)
   if (not valid_pmark and not run_fs) or any((not one_font(f,expected) or sz!='24') for _,f,sz in run_fs):
    fails.append((i+1,expected,fs))
print('FAILURE_COUNT',len(fails))
for x in fails[:80]: print('FAIL',x)
if fails: raise SystemExit(1)
print('RESULT: 通过')
