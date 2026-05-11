import subprocess, sys
from pathlib import Path

ROOT=Path(__file__).resolve().parent

def run(name,*args):
    cmd=[sys.executable,str(ROOT/name),*map(str,args)]
    print('\n## RUN',' '.join(cmd),flush=True)
    cp=subprocess.run(cmd)
    print('## EXIT',name,cp.returncode,flush=True)
    return cp.returncode

def main():
    if len(sys.argv)<3:
        print('USAGE: verify_all_sections_v1.py <source.docx> <output.docx>')
        return 2
    src=Path(sys.argv[1]); out=Path(sys.argv[2])
    failures=[]
    checks=[
        ('audit_document_full_v1.py',out),
        # TOC is protected by default: preserve source TOC instead of rebuilding/reformatting it.
        ('verify_toc_preserved_v1.py',src,out),
        ('verify_headings_strict_v1.py',out),
        ('verify_headings_against_source_v1.py',src,out),
        ('verify_skill_output_v1.py',out),
        ('audit_formatting_coverage_v1.py',src,out),
        ('diff_docx_text_v1.py',src,out),
        ('audit_docx_integrity_v1.py',out),
        ('inspect_media_tables.py',out),
        ('verify_blank_single_font_v1.py',out),
        ('verify_strict_indent_blank_v1.py',out),
    ]
    for item in checks:
        code=run(item[0],*item[1:])
        if code!=0:
            failures.append((item[0],code))
    print('\nVERIFY_ALL_FAILURE_COUNT',len(failures))
    for name,code in failures:
        print('FAIL_MODULE',name,'exit',code)
    print('RESULT','通过' if not failures else '未通过')
    return 0 if not failures else 1
if __name__=='__main__': raise SystemExit(main())
