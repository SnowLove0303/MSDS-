# Agent input contract

The Agent receives a labeled source-facts object and explicit output choices. The
source-facts object is not the same as a formal `write_items` object: it is retained
for traceability and must include the source location for every input fact.

Minimum source-facts fields:

```json
{
  "source": {"path": "...", "sha256": "...", "sheet": "..."},
  "product": {"model": "PEA-4139", "language": "zh", "brand": "guanzhi"},
  "sections": {
    "1": [{"label": "中文名称", "value": "...", "source_ref": "Sheet1!B2"}],
    "3": {"产品类型": "混合物", "components": [{"name": "...", "cas": "...", "conc": "...", "source_ref": "Sheet1!B8"}]},
    "9": [{"label": "外观", "value": "...", "source_ref": "Sheet1!B20"}]
  }
}
```

The formal payload adds derived sections and an evidence ledger. Every derived value
must be traceable to one or more evidence records. Preserve raw source values in a
separate field even when a normalized label or unit is used for mapping.

Allowed output choices:

- `brand`: `guanzhi` or `guocai`
- `language`: `zh` or `en`
- `output_format`: `docx` or `pdf`
- `template_path`: one of the two engine-approved templates only

Do not put `无数据`/`不适用` into the formal values. If the source itself contains
one of those strings, keep it in the raw-facts audit record and omit it from the
formal value channel. S11/S12 use the engine fallback policy instead.
