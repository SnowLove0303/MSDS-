---
name: msds-overwrite-agent
description: >
  Run the Guanzhi MSDS retrieval, inference, and template-driven overwrite workflow.
  Use when an Agent receives labeled Section 1, Section 3, and Section 9 facts and must
  search evidence, map the database-canonical 17-section skeleton, and produce a
  controlled Word/PDF result for a selected brand and language. Do not overwrite a
  source workbook or an approved template.
---

# MSDS retrieval and template overwrite

This skill is the Agent-facing production path built on the local Guanzhi MSDS modules.
It is a controlled document transformation workflow, not permission to invent a
chemical fact or to claim legal certification.

## Fixed contract

- The internal model is the database-canonical S0 + S1–S16 skeleton. Section 9 is
  always represented internally by all 37 standard fields, including 9.37 `其他信息`.
- The Agent's source entry point is labeled S1, S3, and S9 data. Copy the user's
  values exactly into the raw-facts record before normalization. Preserve units,
  comparison symbols, ranges, spelling, and `商业机密` text.
- S3 is a product type plus a component array with `name`, `cas`, and `conc`.
- Every derived field must have an evidence record: query, source URL or local
  retrieval record, retrieval date, jurisdiction/standard version, and the reason
  the source supports the field. A search result is not itself a conclusion.
- The final Word/PDF must be produced by the approved template overwrite engine.
  Never edit the user's source file or use an arbitrary Word file as the container.
  The approved templates are the CN and EN PEA-4139 Guanzhi templates enforced by
  `05-覆写模块/msds_overwrite_engine.py`.
- Brand, language, and output format are explicit inputs (`guanzhi|guocai`,
  `zh|en`, `docx|pdf`). Do not infer a brand from a filename.

## Required execution order

1. Read the source workbook/document in read-only mode. Record path, SHA-256,
   sheet/paragraph coordinates, and an immutable raw input snapshot.
2. Classify each source value into the standard skeleton using the retrieval system
   in `F:\正式项目与模块化内容\冠志\MSDS\02-检索系统`. Do not write to the source.
3. Build the database-standard 17-section input/derivation object. Keep source facts,
   derived facts, evidence, conflicts, and manual-review items in separate channels.
4. Search authoritative sources and applicable Chinese regulations/standards. At
   minimum, check the current official status/version for the claimed GB/T 16483,
   GB 30000 classification requirements, GB 15258 labeling requirements, and any
   product/component-specific regulatory source. Record a source as current only
   after checking its issuing authority and effective status.
5. Derive S2 and S4–S16 only from S1/S3/S9 plus evidence. If evidence conflicts with
   a user's source fact, preserve the source fact and create a conflict for review;
   do not silently replace it.
6. Validate the Agent payload with `scripts/validate_overwrite_input.py`, then call
   the overwrite module using an approved template path. Keep outputs in an isolated
   output directory and retain the write-items JSON beside the draft.
7. Read the produced Word back with the retrieval reader. Confirm section ownership,
   field mapping, S9 presentation, S11/S12 evidence policy, S15/S16 updates, and
   forbidden-placeholder scan before offering a download.

## Formal presentation rules

- Empty S9 fields are internal schema fields but are not rendered in formal Word/PDF.
  Only S9 fields with a real value are emitted. Do not convert blank to `无数据` or
  `不适用`.
- Do not present exact missing/not-applicable placeholders in the formal document.
  The input/audit package may retain an original placeholder as a source fact, but
  the formal output gate must filter it.
- If no component/solvent/additive toxicology research is found for S11, output only
  this one line in the S11 body: `该产品无可用的毒理学研究。`
- If no ecological research is found for S12, output only this controlled analogous
  line: `该产品无可用的生态毒理学研究。`
- Do not retain old template research rows when the current evidence set is absent.
  The engine rebuilds S11/S12 from the current evidence set or the controlled single
  fallback line.
- S15 exposes two optional explanation slots plus the law-item list. Keep the
  section structure and write user content below the fixed headings. S16 exposes the
  disclaimer note as an editable slot. A changed value must be included in the
  payload, not merely displayed in a GUI.
- English translation is a work product that still requires human review. A machine
  translation or brand substitution is not evidence of legal conformity.

## Payload shape

Use the standard overwrite payload after evidence derivation:

```json
{
  "sections": {
    "1": [{"seq": "1.1", "label": "中文名称", "value": "PEA-4139"}],
    "3": {"产品类型": "混合物", "components": [{"name": "...", "cas": "...", "conc": "..."}]},
    "9": [{"seq": "9.1", "label": "外观", "value": "..."}],
    "11": [{"label": "急性毒性", "value": "evidence-backed text"}],
    "12": [{"label": "生态毒性", "value": "evidence-backed text"}],
    "15": [
      {"label": "其它的规定", "value": "optional reviewed note"},
      {"label": "符合下列法规要求", "value": "optional reviewed note"},
      {"label": "法规条目", "value": "GB/T 16483-2008"}
    ],
    "16": [{"label": "免责声明", "value": "reviewed disclaimer"}]
  },
  "keep_structure": "all",
  "empty_policy": "preserve",
  "missing_policy": "preserve",
  "missing_text": "",
  "formal_output_policy": {
    "suppress_placeholders": true,
    "show_s9_only_with_value": true,
    "evidence_fallback": true
  }
}
```

Before the formal call, ensure S1/S3/S9 exist in the source facts even if the
current derivation payload also contains additional sections. A missing S11/S12
section is interpreted by the engine as “no current evidence” and yields the
controlled single-line fallback; an explicitly searched section should instead carry
its evidence-backed fields.

## Failure and handoff rules

- Stop if the template is not one of the two approved templates, if a source/template
  hash changes unexpectedly, if a field cannot be assigned to the standard skeleton,
  or if a legal claim lacks an evidence record.
- Stop formal delivery if readback fails, if S9 contains placeholder/empty rows, if
  S11/S12 contain stale template rows, or if S15/S16 values do not round-trip.
- Report `formal_ready=false` when any translation review, evidence conflict, or
  structural audit remains. State exactly which sections and fields need review.
- Never report “compliant”, “legally certified”, or “verified” merely because a Word
  file was created. Report “template-rendered draft; human/regulatory review pending”
  unless the responsible reviewer has separately approved it.

## Local implementation references

- Retrieval/classification: `F:\正式项目与模块化内容\冠志\MSDS\02-检索系统`
- Inference/evidence: `F:\正式项目与模块化内容\冠志\MSDS\04-推断引擎`
- Approved overwrite engine: `F:\正式项目与模块化内容\冠志\MSDS\05-覆写模块`
- Active Web GUI/API: `F:\正式项目与模块化内容\冠志\MSDS\08-正式版程序`
- Input contract validator: `scripts/validate_overwrite_input.py`
- Database skeleton reference: `../msds-inference-skeleton/references/standard_skeleton.json`
