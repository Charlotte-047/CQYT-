from zipfile import ZipFile
from pathlib import Path
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
def style(p):
 ppr=p.find(q('pPr')); ps=None if ppr is None else ppr.find(q('pStyle'))
 return '' if ps is None else ps.get(q('val')) or ''
def ind(p):
 ppr=p.find(q('pPr')); i=None if ppr is None else ppr.find(q('ind'))
 return {} if i is None else {k.split('}')[-1]:v for k,v in i.attrib.items()}
def blank_expected_font(paras, idx):
 # English abstract/front-matter blanks use Times New Roman; Chinese/body blanks use 宋体.
 state='front'
 for j in range(0, idx+1):
  tt=text(paras[j])
  if tt=='ABSTRACT': state='abstract_en'
  elif norm(tt)==(chr(0x76ee)+chr(0x5f55)): state='toc'
  elif re.match(r'^1\s+\S+', tt or ''): state='body'
 return 'Times New Roman' if state=='abstract_en' else '宋体'

def rpr_fonts(p):
 out=[]
 ppr=p.find(q('pPr'))
 if ppr is not None:
  rpr=ppr.find(q('rPr'))
  if rpr is not None: out.append(('pMark',rpr))
 for r in p.findall('./w:r',NS):
  rpr=r.find(q('rPr'))
  if rpr is not None: out.append(('run',rpr))
 vals=[]
 for kind,rpr in out:
  rf=rpr.find(q('rFonts')); sz=rpr.find(q('sz'))
  vals.append((kind, None if rf is None else {k.split('}')[-1]:v for k,v in rf.attrib.items()}, None if sz is None else sz.get(q('val'))))
 return vals
def one_font(fdict,name):
 if not fdict: return False
 if name=='Times New Roman':
  return all(fdict.get(k)=='Times New Roman' for k in ('eastAsia','ascii','hAnsi','cs'))
 if name=='宋体':
  all_song=all(fdict.get(k)=='宋体' for k in ('eastAsia','ascii','hAnsi','cs'))
  mixed=fdict.get('eastAsia')=='宋体' and fdict.get('ascii')=='Times New Roman' and fdict.get('hAnsi')=='Times New Roman'
  return all_song or mixed
 return all(fdict.get(k)==name for k in ('eastAsia','ascii','hAnsi','cs'))
def is_body_heading_text(t):
 if not t or len(t)>90 or re.search(r'[。；;]$',t): return False
 return bool(re.match(r'^\d+(?:\.\d+){0,2}\s+\S.{0,80}$',t)) or norm(t) in (chr(0x81f4)+chr(0x8c22), chr(0x53c2)+chr(0x8003)+chr(0x6587)+chr(0x732e))

def is_body_text(t): return t and not re.match(r'^(\[\d+\]|图\d+\.\d+|表\d+\.\d+|注：)',t)
p=Path(sys.argv[1]); fails=[]
with ZipFile(p) as z:
 root=ET.fromstring(z.read('word/document.xml'))
 paras=root.findall('.//w:p',NS)
 abstract_idx=None
 for i,pa in enumerate(paras):
  if text(pa).upper()=='ABSTRACT': abstract_idx=i
 state='front'
 for i,n in enumerate(paras):
  t=text(n); st=style(n)
  if t=='1 绪论': state='body'
  if t=='参考文献': state='refs'
  if n.getparent().tag==q('body') and (raw(n)==' ' or t=='') and not has_drawing(n):
   expected=blank_expected_font(paras, i)
   fonts=rpr_fonts(n)
   valid_pmark=any(x[0].lower()=='pmark' and (one_font(x[1],expected) or (expected==chr(0x5b8b)+chr(0x4f53) and (x[1].get('eastAsia')==chr(0x5b8b)+chr(0x4f53) or x[1].get('cs')==chr(0x5b8b)+chr(0x4f53)))) and (x[2] in ('24', None)) for x in fonts)
   run_fonts=[x for x in fonts if x[0].lower()=='run' and x[2] is not None] or ([] if valid_pmark else fonts)
   if (not valid_pmark and not run_fonts) or any((not one_font(f,expected) or sz!='24') for _,f,sz in run_fonts):
    fails.append(f'blank font/size wrong near paragraph {i+1}; expected {expected}: {fonts}')
  if n.getparent().tag==q('body') and state=='body' and not is_body_heading_text(t) and st not in ('2','3','4','5','Heading1','Heading2','Heading3','Heading4') and is_body_text(t):
   if t.startswith('参考文献') or re.match(r'^\[\d+\]',t): continue
   ia=ind(n)
   if ia.get('firstLineChars')!='200' or 'firstLine' in ia:
    fails.append(f'body indent not strict 2 chars: {t[:30]} {ia}')
print('FAILURE_COUNT',len(fails))
for f in fails[:80]: print('FAIL',f)
if fails: raise SystemExit(1)
print('RESULT: 通过')
