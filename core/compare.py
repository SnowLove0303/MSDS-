# -*- coding: utf-8 -*-
"""模板覆写指向分析.

输入: 模板 ParseResult + 企业产/新产品 ParseResult.
输出: 逐字段的覆写指向 (write_source), 用于判断:
  - 模板字段应被产品值覆盖 / 保留模板值 / 清空 / 新增
  - 核心输入集 A (S1/S3/S9) 与外推集 B (S2/S4-16) 分开处理
"""
from __future__ import annotations

from dataclasses import dataclass, field

from .structure import ParseResult

# 输入集 A: 用户/源文档明确提供的事实字段
# 外推集 B: 由 A + 规则库推导
INPUT_SET_A = (1, 3, 9)
OUTPUT_SET_B = tuple(n for n in range(1, 17) if n not in (1, 3, 9))


@dataclass
class OverwriteDecision:
    """一个字段的覆写指向."""
    section: int
    label: str
    template_value: str        # 模板现值
    product_value: str         # 产品提供值
    write_source: str          # "template" | "product" | "clear" | "add" | "review"
    reason: str = ""
    matched: bool = False      # 产品是否有该字段


@dataclass
class CompareResult:
    """模板 vs 产品的整体比对结果."""
    template_file: str = ""
    product_file: str = ""
    decisions: list[OverwriteDecision] = field(default_factory=list)
    # 按节统计
    by_section: dict[int, dict[str, int]] = field(default_factory=dict)
    # 异常标记: 模板有值但产品没提供 → 可能残留旧值
    residue_risks: list[OverwriteDecision] = field(default_factory=list)
    # 核心输入集 A 覆盖情况
    set_a_coverage: dict[int, dict] = field(default_factory=dict)

    def counts(self) -> dict[str, int]:
        from collections import Counter
        return dict(Counter(d.write_source for d in self.decisions))


def compare(template: ParseResult, product: ParseResult) -> CompareResult:
    """比对模板与产品, 生成覆写指向."""
    result = CompareResult(
        template_file=template.file_name,
        product_file=product.file_name,
    )

    # 产品字段全集 (label → value), 含成分表转成字段
    # section 0 (页眉页脚) 不参与自动覆写比对: 其字段编辑状态由手动标注决定
    product_fields: dict[tuple[int, str], str] = {}
    for num, sec in product.sections.items():
        if num == 0:
            continue
        for f in sec.fields:
            product_fields[(num, f.label)] = f.value

    # 模板字段全集
    template_fields: dict[tuple[int, str], str] = {}
    for num, sec in template.sections.items():
        if num == 0:
            continue
        for f in sec.fields:
            template_fields[(num, f.label)] = f.value

    # 合并键 (模板字段优先遍历)
    all_keys = set(template_fields) | set(product_fields)

    for key in sorted(all_keys, key=lambda k: (k[0], k[1])):
        num, label = key
        tval = template_fields.get(key, "")
        pval = product_fields.get(key, "")

        if key in template_fields and key in product_fields:
            # 模板和产品都有
            if tval == pval:
                write = "template"   # 值一致, 保留模板即可
                reason = "模板与产品值一致"
            elif tval and not pval:
                write = "review"     # 不应发生 (双方都有但产品空)
                reason = "产品字段为空, 模板有旧值"
            else:
                write = "product"    # 产品提供了不同值 → 用产品覆盖
                reason = "产品值覆盖模板值"
        elif key in template_fields:
            # 仅模板有 → 产品没提供
            if tval:
                write = "review"     # 模板残留旧值风险
                reason = "模板有旧值, 产品未提供 → 需确认是否保留"
            else:
                write = "clear"      # 模板空字段, 产品也没填 → 保持空
                reason = "模板字段未填, 产品未提供"
        else:
            # 仅产品有 → 新增字段
            write = "add"
            reason = "产品新增字段"

        dec = OverwriteDecision(
            section=num, label=label,
            template_value=tval, product_value=pval,
            write_source=write, reason=reason,
            matched=key in product_fields,
        )
        result.decisions.append(dec)

        # 按节统计
        sec_stat = result.by_section.setdefault(num, {"template": 0, "product": 0, "clear": 0, "add": 0, "review": 0})
        sec_stat[write] = sec_stat.get(write, 0) + 1

        if write == "review" and tval:
            result.residue_risks.append(dec)

    # 输入集 A 覆盖情况 (S1/S3/S9)
    for num in INPUT_SET_A:
        tsec = template.sections.get(num)
        psec = product.sections.get(num)
        coverage = {
            "template_fields": len(tsec.fields) if tsec else 0,
            "product_fields": len(psec.fields) if psec else 0,
            "components_template": len(tsec.components) if tsec else 0,
            "components_product": len(psec.components) if psec else 0,
        }
        # 成分逐项比对
        if tsec and psec:
            coverage["component_match"] = _compare_components(tsec.components, psec.components)
        result.set_a_coverage[num] = coverage

    return result


def _compare_components(tpl_components, prod_components) -> dict:
    """成分表比对: 名称/CAS 逐项."""
    t_names = {(c.name, c.cas) for c in tpl_components}
    p_names = {(c.name, c.cas) for c in prod_components}
    return {
        "matched": len(t_names & p_names),
        "template_only": len(t_names - p_names),
        "product_only": len(p_names - t_names),
        "total": len(t_names),
    }
