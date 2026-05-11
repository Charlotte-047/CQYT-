from pathlib import Path
import argparse, shutil, subprocess, sys, time

ROOT = Path(__file__).resolve().parent

def run(script, *args, check=True):
    cmd = [sys.executable, str(ROOT / script), *map(str, args)]
    print('$', ' '.join(cmd), flush=True)
    cp = subprocess.run(cmd)
    if check and cp.returncode != 0:
        raise SystemExit(cp.returncode)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('input')
    ap.add_argument('output')
    ns = ap.parse_args()

    src = Path(ns.input).resolve()
    out = Path(ns.output).resolve()
    stamp = time.strftime('xmlonly-%Y%m%d-%H%M%S')
    work = ROOT.parent / 'pipeline' / f'{stamp}-{src.stem}'
    work.mkdir(parents=True, exist_ok=True)

    step00 = work / '00-prep-step14-input.docx'
    step14 = work / '14-patchv4-format-preserve-media.docx'
    step17 = work / '17-titleblack-pages-h1space.docx'
    step18 = work / '18-media-tables.docx'
    step19 = work / '19-references.docx'
    step20 = work / '20-baseline-layout.docx'
    step21 = work / '21-final-xml-cleanup.docx'

    run('prepare_step14_input_v1.py', src, step00)

    # Reuse existing earlier chain input when available, otherwise assume 06 exists is not required here.
    six = work / '06-heading-numbers-preserved.docx'
    shutil.copyfile(step00, six)
    run('patch_v4_requirements.py', six, step14)
    run('step17_xml_fix_titles_pages_h1space_v2.py', step14, step17)
    run('step18_xml_fix_media_tables_v3.py', step17, step18)
    run('step19_xml_fix_references_v2.py', step18, step19)
    run('step20_xml_fix_baseline_layout.py', step19, step20)
    shutil.copyfile(step20, step21)
    run('final_xml_cleanup_v1.py', step21, check=False)
    shutil.copyfile(step21, out)
    # Defensive final pass on the deliverable itself. Some XML cleanup steps are
    # order-sensitive; the output must be the artifact that gets cleaned/audited.
    run('final_xml_cleanup_v1.py', out, check=False)
    run('restore_source_toc_v1.py', src, out, check=False)
    run('verify_skill_output_v1.py', out, check=False)
    run('verify_headings_strict_v1.py', out, check=False)
    run('verify_headings_against_source_v1.py', src, out, check=False)
    run('verify_all_sections_v1.py', src, out, check=False)
    print('FINAL_OUTPUT:', out)
    print('WORKDIR:', work)

if __name__ == '__main__':
    main()
