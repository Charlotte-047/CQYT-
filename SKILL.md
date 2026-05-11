# paper-thesis-formatter

## Purpose

Format `.docx` thesis documents according to a structured university-thesis style profile and verify that key document content was preserved.

This skill is for reusable thesis-formatting work. It should not contain private thesis files, generated outputs, or user-specific local documents.

## When to use

Use this skill when the user wants to:

- normalize thesis formatting in a `.docx` file;
- apply consistent heading/body/table/front-matter rules;
- check whether formatting preserved headings, tables, images, and long text;
- produce a formatted `.docx` plus an automated verification result.

Do **not** use it as proof that a document is ready for submission without manual review.

## Inputs

Required:

1. source `.docx` thesis file;
2. desired output `.docx` path.

Recommended:

- a school formatting specification;
- a publishable/non-private template or rule profile;
- manual review after output generation.

## Outputs

- formatted `.docx` document;
- verifier console output;
- optional debug/audit reports if the user asks for them.

Generated outputs and user thesis files must not be committed to a public repository.

## Canonical command

```bash
python scripts/format_paper_xml_only.py "input.docx" "output.docx"
python scripts/verify_all_sections_v1.py "input.docx" "output.docx"
```

The output is accepted only if the verifier reports:

```text
VERIFY_ALL_FAILURE_COUNT 0
RESULT 通过
```

## Current formatting strategy

The formatter works at `.docx` XML level.

Current default behavior:

- format H1/H2/H3 headings;
- format body paragraphs;
- format table text and table borders toward a three-line style;
- preserve images/media relationships;
- format detectable abstract/front-matter headings;
- keep H1 headings on new pages;
- preserve existing source TOC by default.

## TOC protection rule

The table of contents is a protected area.

Default rule:

- do not rebuild TOC;
- do not update TOC page numbers;
- do not restyle TOC;
- do not delete or rewrite TOC internals;
- if the source has a TOC, restore/preserve the source TOC block in the output;
- if the source has no detectable TOC, fail and ask for human handling instead of inventing a fake TOC.

Rationale: Word TOC fields and tabs are fragile. Static fake TOCs, Word COM updates, and cloned template TOCs can change page numbers, spacing, fields, or unrelated formatting.

## Required checks

Run the top-level verifier after every formatting run:

```bash
python scripts/verify_all_sections_v1.py "input.docx" "output.docx"
```

Important sub-checks include:

- `verify_toc_preserved_v1.py` — source TOC preservation;
- `verify_headings_against_source_v1.py` — heading count/order preservation;
- `verify_headings_strict_v1.py` — safe heading formatting;
- `audit_formatting_coverage_v1.py` — formatting coverage and content preservation;
- `diff_docx_text_v1.py` — long paragraph loss check;
- `audit_docx_integrity_v1.py` — document integrity/media/page-structure checks;
- `inspect_media_tables.py` — media/table scan;
- `verify_blank_single_font_v1.py` and `verify_strict_indent_blank_v1.py` — blank/indent checks.

A pipeline audit pass alone is not enough; always audit the produced document.

## Privacy rules

Never commit private user documents or generated outputs.

Exclude:

- `.docx`, `.doc`, `.pdf`, `.wps`, `.xlsx`, `.xls`;
- `output/`, `outputs/`, `pipeline/`;
- handover files, local debug reports, memory notes;
- any thesis content or personal data.

## Limitations

- Direct support is for `.docx`; convert `.doc` first.
- Automatic TOC generation/updating is intentionally disabled by default.
- The rules are not universal for every university template.
- Word/WPS/LibreOffice may render XML differently.
- Equations, floating objects, complex fields, and unusual section layouts may need manual review.
- Passing automated checks does not guarantee acceptance by a school or advisor.

## Safety / operating principles

- Always preserve the source file.
- Make all fixes in scripts, not in one-off edited output files.
- Prefer privacy-preserving repository contents.
- Treat automated verification as a gate, not as a substitute for human inspection.
