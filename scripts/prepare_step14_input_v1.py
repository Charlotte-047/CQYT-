from zipfile import ZipFile, ZIP_DEFLATED
from pathlib import Path
from lxml import etree as ET
import sys, tempfile, shutil, re

W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

def q(t): return f'{W}{t}'
def text(p): return ''.join(t.text or '' for t in p.findall('.//w:t',NS)).strip()
def pstyle(p):
    ppr=p.find(q('pPr')); ps=None if ppr is None else ppr.find(q('pStyle'))
    return '' if ps is None else ps.get(q('val')) or ''
def ensure(parent, tag, first=False):
    n=parent.find(q(tag))
    if n is None:
        n=ET.Element(q(tag))
        if first: parent.insert(0,n)
        else: parent.append(n)
    return n

def set_pstyle(p,val):
    ppr=ensure(p,'pPr',True); ps=ensure(ppr,'pStyle',True); ps.set(q('val'),val)

def set_ppr(p,align,first=None,line=440,before=0,after=0):
    ppr=ensure(p,'pPr',True)
    jc=ensure(ppr,'jc'); jc.set(q('val'),align)
    sp=ensure(ppr,'spacing'); sp.set(q('before'),str(before)); sp.set(q('after'),str(after)); sp.set(q('line'),str(line)); sp.set(q('lineRule'),'exact')
    ind=ensure(ppr,'ind')
    for a in ('firstLine','left','hanging','right','firstLineChars','leftChars','hangingChars','rightChars'):
        ind.attrib.pop(q(a),None)
    if first is not None: ind.set(q('firstLine'),str(first))

def set_run(r,east='宋体',asc='Times New Roman',size=24,bold=False,color=None):
    rpr=ensure(r,'rPr',True)
    for c in list(rpr):
        if c.tag in (q('rFonts'),q('sz'),q('szCs'),q('b'),q('bCs'),q('color'),q('u')): rpr.remove(c)
    rf=ET.SubElement(rpr,q('rFonts')); rf.set(q('eastAsia'),east); rf.set(q('ascii'),asc); rf.set(q('hAnsi'),asc); rf.set(q('cs'),asc)
    b=ET.SubElement(rpr,q('b')); b.set(q('val'),'1' if bold else '0')
    bcs=ET.SubElement(rpr,q('bCs')); bcs.set(q('val'),'1' if bold else '0')
    sz=ET.SubElement(rpr,q('sz')); sz.set(q('val'),str(size))
    szcs=ET.SubElement(rpr,q('szCs')); szcs.set(q('val'),str(size))
    if color:
        c=ET.SubElement(rpr,q('color')); c.set(q('val'),color)

def set_all_runs(p,east,asc,size,bold=False,color=None):
    rs=p.findall('./w:r',NS)
    if not rs:
        r=ET.SubElement(p,q('r')); ET.SubElement(r,q('t')).text=''; rs=[r]
    for r in rs: set_run(r,east,asc,size,bold,color)

def normalize(s): return re.sub(r'\s+','',(s or '').strip())
def strip_num(t): return re.sub(r'^\d+(?:\.\d+){0,2}\s*','',t.strip())
def is_toc_line(p): return bool(p.findall('.//w:tab',NS)) or bool(re.match(r'^(\d+(?:\.\d+){0,2}\s+.+?|致谢|参考文献)\s*\d+$', text(p)))
def is_numbered_h1(t): return bool(re.match(r'^\d+\s+\S+',t))
def is_numbered_h2(t): return bool(re.match(r'^\d+\.\d+\s+\S+',t))
def is_numbered_h3(t): return bool(re.match(r'^\d+\.\d+\.\d+\s+\S+',t))

KNOWN = {
 '绪论':'1','课题研究的背景':'1.1','研究背景':'1.1','研究内容':'1.2',
 '制作软件':'2','剪映专业版':'2.1','制作过程':'3','拍摄素材':'3.1','镜头拍摄':'3.1.1','前期拍摄':'3.1.2',
 '素材分类筛选与整理':'3.2','视频前期基础剪辑制作':'3.3','旁白创作':'3.3.1','旁白与画面的配合逻辑':'3.3.2',
 '结论':'4','研究成果总结':'4.1','研究的局限性':'4.2','后续研究展望':'4.3',
 '致谢':'H1','参考文献':'H1'
}
KNOWN_N={normalize(k):v for k,v in KNOWN.items()}

BODY_START_TITLES={normalize('绪论'),normalize('1 绪论')}

def replace_text_keep_one_run(p,new):
    for r in list(p.findall('./w:r',NS)): p.remove(r)
    r=ET.SubElement(p,q('r')); t=ET.SubElement(r,q('t')); t.text=new
    if new.startswith(' ') or new.endswith(' '): t.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
    return r

def classify_body_heading(t, st):
    sk={'3':'Heading1','4':'Heading2','5':'Heading3'}.get(st,st)
    if sk=='Heading1': return 1, None
    if sk=='Heading2': return 2, None
    if sk=='Heading3': return 3, None
    if is_numbered_h3(t): return 3, None
    if is_numbered_h2(t): return 2, None
    if is_numbered_h1(t): return 1, None
    key=normalize(strip_num(t))
    val=KNOWN_N.get(key)
    if val=='H1': return 1, None
    if val:
        return val.count('.')+1, val
    return None, None

def patch_doc(tmp):
    docxml=tmp/'word'/'document.xml'; tree=ET.parse(str(docxml)); root=tree.getroot(); body=root.find(q('body'))
    real=False; counts={1:0,2:0,3:0}; numbered=0
    for p in body.findall('./w:p',NS):
        t=text(p); st=pstyle(p)
        if not real and not is_toc_line(p):
            key=normalize(strip_num(t))
            if key in BODY_START_TITLES or KNOWN_N.get(key): real=True
        if not real or is_toc_line(p): continue
        lvl,num=classify_body_heading(t,st)
        if lvl is None: continue
        title=strip_num(t)
        if num and num!='H1':
            new=f'{num} {title}'
            if text(p)!=new:
                replace_text_keep_one_run(p,new); t=new; numbered+=1
        if lvl==1:
            set_pstyle(p,'3'); set_ppr(p,'center',0); set_all_runs(p,'黑体','Times New Roman',32,False,'000000'); counts[1]+=1
        elif lvl==2:
            set_pstyle(p,'4'); set_ppr(p,'left',0); set_all_runs(p,'黑体','Times New Roman',30,False,'000000'); counts[2]+=1
        elif lvl==3:
            set_pstyle(p,'5'); set_ppr(p,'left',0); set_all_runs(p,'黑体','Times New Roman',28,False,'000000'); counts[3]+=1
    tree.write(str(docxml),xml_declaration=True,encoding='utf-8',standalone='yes')
    return counts, numbered

def patch_styles(tmp):
    styles=tmp/'word'/'styles.xml'
    if not styles.exists(): return
    tree=ET.parse(str(styles)); root=tree.getroot()
    for sid,size,align in [('3','32','center'),('4','30','left'),('5','28','left'),('Heading1','32','center'),('Heading2','30','left'),('Heading3','28','left')]:
        st=root.find(f".//w:style[@w:styleId='{sid}']",NS)
        if st is None: continue
        ppr=ensure(st,'pPr'); jc=ensure(ppr,'jc'); jc.set(q('val'),align)
        ind=ensure(ppr,'ind')
        for a in list(ind.attrib): ind.attrib.pop(a,None)
        ind.set(q('firstLine'),'0')
        rpr=ensure(st,'rPr')
        for c in list(rpr):
            if c.tag in (q('rFonts'),q('sz'),q('szCs'),q('b'),q('bCs'),q('color'),q('u')): rpr.remove(c)
        rf=ET.SubElement(rpr,q('rFonts')); rf.set(q('eastAsia'),'黑体'); rf.set(q('ascii'),'Times New Roman'); rf.set(q('hAnsi'),'Times New Roman'); rf.set(q('cs'),'Times New Roman')
        b=ET.SubElement(rpr,q('b')); b.set(q('val'),'0')
        bcs=ET.SubElement(rpr,q('bCs')); bcs.set(q('val'),'0')
        sz=ET.SubElement(rpr,q('sz')); sz.set(q('val'),size)
        szcs=ET.SubElement(rpr,q('szCs')); szcs.set(q('val'),size)
        color=ET.SubElement(rpr,q('color')); color.set(q('val'),'000000')
    tree.write(str(styles),xml_declaration=True,encoding='utf-8',standalone='yes')

def process(src,out):
    src=Path(src); out=Path(out); out.parent.mkdir(parents=True,exist_ok=True)
    tmp=Path(tempfile.mkdtemp(prefix='prep_step14_'))
    try:
        with ZipFile(src) as z: z.extractall(tmp)
        counts,numbered=patch_doc(tmp); patch_styles(tmp)
        repack=out.with_suffix('.tmp.docx')
        with ZipFile(repack,'w',ZIP_DEFLATED) as z:
            for f in tmp.rglob('*'):
                if f.is_file(): z.write(f,f.relative_to(tmp))
        shutil.move(str(repack),str(out))
        print('OUTPUT:',out); print('HEADING_COUNTS_PREP:',counts); print('NUMBERED:',numbered)
    finally: shutil.rmtree(tmp,ignore_errors=True)

if __name__=='__main__':
    if len(sys.argv)!=3: raise SystemExit('usage: python prepare_step14_input_v1.py input.docx output.docx')
    process(sys.argv[1],sys.argv[2])
