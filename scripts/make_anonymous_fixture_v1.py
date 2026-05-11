from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.sax.saxutils import escape
import sys

W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
CT='http://schemas.openxmlformats.org/package/2006/content-types'

def p(text='', style=None, pagebreak=False):
    ppr=''
    if style or pagebreak:
        parts=[]
        if pagebreak: parts.append('<w:pageBreakBefore/>')
        if style: parts.append(f'<w:pStyle w:val="{style}"/>')
        ppr='<w:pPr>'+''.join(parts)+'</w:pPr>'
    return f'<w:p>{ppr}<w:r><w:t>{escape(text)}</w:t></w:r></w:p>'

def make(path):
    body=[]
    body.append(p('原创性声明'))
    body.append(p('摘  要'))
    body.append(p('这是一段匿名中文摘要，用于测试格式化流程，不包含任何私人论文内容。'))
    body.append(p('关键词：匿名；测试；格式'))
    body.append(p('ABSTRACT'))
    body.append(p('This is an anonymous abstract paragraph for pipeline testing only.'))
    body.append(p('Key words: anonymous; test; format'))
    body.append(p('目    录'))
    body.append(p('1 Introduction 1', 'TOC1'))
    body.append(p('1.1 Background 1', 'TOC2'))
    body.append(p('2 Method 2', 'TOC1'))
    body.append(p('1 绪论', '3', True))
    body.append(p('这是一段匿名正文内容，用于验证正文格式化。它足够长，包含中文句号。'))
    body.append(p('1.1 研究背景', '4'))
    body.append(p('这是一段二级标题下的匿名正文内容，用于验证标题层级和正文格式。'))
    body.append(p('1.1.1 研究细节', '5'))
    body.append(p('这是一段三级标题下的匿名正文内容，用于验证格式化稳定性。'))
    body.append(p('2 方法', '3', True))
    body.append(p('这是一段第二章匿名正文内容，用于验证分页和章节格式。'))
    body.append(p('参考文献', '3', True))
    body.append(p('[1] Anonymous Author. Anonymous Reference Title[J]. Journal, 2026.'))
    sect='<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="907" w:footer="992" w:gutter="0"/></w:sectPr>'
    document=f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>{''.join(body)}{sect}</w:body></w:document>'''
    styles=f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W}">
<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="3"><w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="4"><w:name w:val="Heading 2"/><w:basedOn w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="5"><w:name w:val="Heading 3"/><w:basedOn w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="TOC 1"/><w:basedOn w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="TOC2"><w:name w:val="TOC 2"/><w:basedOn w:val="Normal"/></w:style>
</w:styles>'''
    rels='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
    docrels='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'''
    ctypes='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'''
    path=Path(path); path.parent.mkdir(parents=True,exist_ok=True)
    with ZipFile(path,'w',ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml',ctypes)
        z.writestr('_rels/.rels',rels)
        z.writestr('word/document.xml',document)
        z.writestr('word/styles.xml',styles)
        z.writestr('word/_rels/document.xml.rels',docrels)
    print(path)
if __name__=='__main__': make(sys.argv[1])
