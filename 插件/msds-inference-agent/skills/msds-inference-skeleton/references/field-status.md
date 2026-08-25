# Facts-package status semantics

The facts package is not a formal MSDS. Every field must carry an origin/status that allows a reviewer to distinguish copied data from requirements and inference.

| status | meaning | allowed value behavior |
|---|---|---|
| `source_fact` | Copied from the submitted workbook | Keep the original value exactly; retain cell coordinates and raw label |
| `derived_fact` | Deterministic derivation from a source fact, such as an S0 model code | Show the derivation basis; never present it as an original cell value |
| `source_not_provided` | The active skeleton requires the field, but S1/S3/S9 did not provide it | Render “原始输入未提供”; do not use a template default |
| `needs_search` | A current authoritative source is needed before the field can be evaluated | Render “需检索”; attach evidence separately when found |
| `manual_review` | Evidence, range, identity, or applicability is insufficient or conflicting | Render “需人工审核”; retain all conflicting values |
| `conflict` | Source and external evidence disagree, or active skeleton baselines disagree | Preserve both sides and explain the conflict |
| `not_applicable` | Applicability is established by a rule and evidence | Record the rule/evidence; do not infer from an empty cell alone |

Search evidence is metadata, not a license to rewrite a `source_fact`. Formal SDS generation is a later, separately authorized workflow.
