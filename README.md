# Paper Thesis Formatter

A Python-based `.docx` thesis formatting toolkit for normalizing university thesis documents and running automated acceptance checks.

> Current rule profile: thesis-style documents similar to Chongqing Yitong College graduation thesis formatting. The project is designed as a reusable formatting pipeline, not as a one-off document edit.

## What it does

The formatter works directly on WordprocessingML inside `.docx` files. It can:

- format first/second/third-level headings;
- normalize body text font, alignment, line spacing, and first-line indent;
- format Chinese/English abstract headings and front-matter sections where detectable;
- format tables toward a three-line-table style;
- preserve images and media relationships;
- keep existing table/image counts stable;
- preserve an existing source table of contents (TOC) by default;
- run a multi-step verifier after formatting.

## Important TOC policy

The TOC is treated as a protected area by default.

The formatter does **not** rebuild, update, or restyle the TOC by default because Word TOC fields are fragile and automatic updates can unexpectedly change layout, page numbers, or other document formatting.

Current default behavior:

- if the source document already contains a TOC, the source TOC block is restored into the output document;
- TOC text, page numbers, tabs, hidden fields, and internal structure are preserved as much as possible;
- if the source document has no detectable TOC, the formatter should fail with a human-action-needed status instead of inventing a fake TOC.

If you need automatic TOC generation or page-number updates, implement it as an explicit opt-in feature and inspect the result manually in Microsoft Word.

## Requirements

- Python 3.10+
- `lxml`
- A `.docx` input file

Install dependency:

```bash
pip install lxml
```

Windows is the primary tested environment. The XML-only pipeline should be mostly platform-independent, but Word rendering still depends on Microsoft Word/WPS/LibreOffice behavior.

## Quick start

```bash
python scripts/format_paper_xml_only.py "input.docx" "output.docx"
python scripts/verify_all_sections_v1.py "input.docx" "output.docx"
```

If the verifier prints:

```text
VERIFY_ALL_FAILURE_COUNT 0
RESULT 通过
```

then the output passed the current automated checks.

## Recommended workflow

1. Make a backup of your thesis document.
2. Ensure the input is `.docx`.
3. Run the formatter.
4. Run the full verifier.
5. Open the output in Word and manually inspect:
   - TOC appearance and page numbers;
   - abstract/front matter;
   - headings and pagination;
   - tables;
   - images and captions;
   - headers/footers/page numbers.
6. Only submit after manual review.

## Acceptance checks

Main gate:

```bash
python scripts/verify_all_sections_v1.py "input.docx" "output.docx"
```

The full verifier calls multiple checks, including structure, heading consistency, formatting coverage, text preservation, TOC preservation, media/table inspection, blank paragraph checks, and document integrity checks.

Notable helper checks include:

```bash
python scripts/verify_headings_against_source_v1.py "input.docx" "output.docx"
python scripts/verify_toc_preserved_v1.py "input.docx" "output.docx"
python scripts/audit_formatting_coverage_v1.py "input.docx" "output.docx"
python scripts/diff_docx_text_v1.py "input.docx" "output.docx"
python scripts/audit_docx_integrity_v1.py "output.docx"
```

## Privacy and test data

This repository intentionally does **not** include private thesis files, generated outputs, or local test documents.

Do not commit:

- `.docx`, `.doc`, `.pdf`, `.wps` files containing thesis content;
- generated `output/` or `pipeline/` folders;
- handover notes or private debugging logs;
- personal templates unless you have permission to publish them.

The `.gitignore` is configured to exclude common private/generated document files.

## Limitations

This project is not a replacement for human review.

Known limitations:

- only `.docx` is supported directly;
- TOC generation/updating is not automatic by default;
- school-specific templates may require additional rule profiles;
- complex fields, section breaks, floating objects, equations, and legacy Word structures can behave differently across Word/WPS/LibreOffice;
- automated checks are conservative but cannot guarantee visual perfection;
- generated output must still be opened and reviewed manually.

## Disclaimer

Use this tool at your own risk. It modifies Word document structure programmatically and may not satisfy every school, advisor, or department-specific formatting requirement. Always keep backups and manually inspect the final document before submission.

The maintainers are not responsible for missed deadlines, formatting rejection, data loss, or incorrect submission caused by using this tool.

## Repository layout

```text
paper-thesis-formatter/
  scripts/      formatter and verifier scripts
  references/   non-private reference notes, if any
  templates/    placeholder for publishable templates only
  README.md
  SKILL.md
```

## Development notes

- Treat source thesis content as private.
- Keep fixes in scripts, not in one-off output files.
- Run `verify_all_sections_v1.py` after every formatter change.
- Prefer XML-only transformations unless an opt-in Word automation step is explicitly required.
- Avoid writing non-ASCII source literals through a misconfigured Windows shell; use UTF-8 files or Unicode-safe constants where needed.
