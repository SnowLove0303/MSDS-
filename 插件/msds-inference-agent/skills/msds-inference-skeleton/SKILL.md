---
name: msds-inference-skeleton
description: >
  Build a source-preserving, evidence-traceable MSDS facts package from Section 1,
  Section 3, and Section 9 Excel inputs and map it to the database-canonical
  S0+S1–S16 standard skeleton. Use when the user needs a first-pass MSDS inference workflow,
  a raw-input facts report, or a standards-audited skeleton before formal MSDS
  generation. Do not silently change source facts or produce a formal MSDS.
---

# MSDS standard-skeleton facts

This skill is the first local implementation of the Guanzhi MSDS inference plugin. It separates three things that must not be conflated:

1. `source_fact`: a value copied from the user's input, with its original cell/label and exact text preserved;
2. `derived_fact`: a deterministic value such as the S0 model extracted from a source value, with its basis recorded;
3. `required_field`: a field required by the active 17-section skeleton but absent from the input. It is marked `source_not_provided`, `needs_search`, or `manual_review`; it is not a formal SDS conclusion.

## Scope and hard boundaries

- The canonical skeleton is the `schema_field` table in `03-数据库/正式库/Data Base/msds_standard.db`: S0 header/footer plus S1–S16, with S9 fixed at 37 database fields and S11 fixed at 10 database fields. S1, S3, and S9 are copied from the input; the other fourteen content sections are represented as standard fields with explicit evidence status.
- The active PEA-4139 Word template is an output-layout baseline only. It has a shorter S9 table and must not replace the database-canonical 37-field S9 skeleton.
- Preserve source text, comparison symbols, units, spacing, numbering anomalies, and “商业机密” exactly. Add normalized labels only as mapping metadata.
- Use official standards, laws, and regulator/database pages to validate the skeleton and record evidence. Search results may establish a source or a field requirement, but must not overwrite a source fact.
- If an input value conflicts with a search result, retain the input value and create a `conflict` or `manual_review` record.
- Missing information must not be filled from product-name intuition, a neighboring product, a previous MSDS, or a generic template default.
- This skill produces a raw-input facts package and a non-formal Word report by default. It must not overwrite the source workbook, an existing Word template, the formal model database, or publish a formal MSDS unless a later request explicitly authorizes that separate workflow.

## Required workflow

1. Read the workbook in read-only mode and record path, SHA-256, workbook sheets, and source coordinates.
2. Parse only the S1, S3, and S9 blocks into exact source facts. Keep inequalities such as `> 40`, `<3`, and `47±2%` unchanged.
3. Load [the database-canonical skeleton reference](references/standard_skeleton.json) and emit all 17 sections, including all 37 S9 fields and all 10 S11 fields.
4. Add standards/regulatory evidence as a separate evidence list. For current or legally sensitive claims, prefer an official government or standards source and record retrieval date and source status.
5. Render a facts-only Word report. Title it as a facts/traceability report and state that it is not a formal MSDS.
6. Validate: all 17 sections exist, source facts round-trip, missing fields are visibly marked, evidence is separate from facts, no formal-MSDS marker is emitted as a source fact, and the input/template hashes are unchanged.

## Initial CLI

```powershell
python scripts/build_facts_skeleton.py `
  --input "<PEA-4139 MSDS表单.xlsx>" `
  --evidence "<regulatory-evidence.json>" `
  --out "<raw_input_facts.json>"

python scripts/render_facts_docx.py `
  --facts "<raw_input_facts.json>" `
  --out "<raw-input-facts-17-section.docx>"
```

The renderer intentionally labels absent values as “原始输入未提供” or “需人工审核”; those labels describe the facts package and are not replacement text for a formal SDS.

For database field names and section counts, read [standard_skeleton.json](references/standard_skeleton.json). For status semantics and evidence separation, read [field-status.md](references/field-status.md). The active PEA-4139 Word template is layout evidence, not the canonical field-count source.
