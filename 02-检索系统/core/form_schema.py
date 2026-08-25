# -*- coding: utf-8 -*-
"""表单系统输入结构 (参照 PEA-4139 MSDS 表单).

表单系统收集「使用者已知的源头数据」——Section 1 / 3 / 9 的可填字段。
输出是这些字段的结构化定义, 供:
  - GUI 表单窗口渲染输入控件 (gui/form_window.py)
  - 后续推断引擎 (生成 S2/S4~16) 与 覆写引擎 (写入项) 共用同一字段锚点

字段标签以 PEA-4139 MSDS表单.xlsx 为准, 与覆写引擎 field_maps.json 的目标标签对齐。
"""
from __future__ import annotations

from dataclasses import dataclass, field, asdict
from typing import Any


@dataclass
class FormField:
    """表单一个可填字段."""
    label: str                 # 字段标签 (模板/表单标准写法)
    kind: str = "text"         # text | multiline | select
    seq: str = ""              # 模板序号 (如 '9.1'); 无序号留空
    placeholder: str = ""      # 输入占位提示
    options: list[str] = field(default_factory=list)   # select 的可选值
    required: bool = False     # 是否必填 (S3 成分必需)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class ComponentRow:
    """S3 成分表一行 (使用者填写)."""
    name: str = ""
    cas: str = ""
    conc: str = ""


@dataclass
class FormSections:
    """表单系统的输入结构: S1 / S3 / S9 三节."""
    s1: list[FormField] = field(default_factory=list)
    s3: list[FormField] = field(default_factory=list)
    s9: list[FormField] = field(default_factory=list)


# ---------- PEA-4139 表单输入结构 ----------

# S1 物料及供应商标识: 使用者填产品/供应商信息
S1_FIELDS: list[FormField] = [
    FormField("产品名称", kind="text", seq="1.1",
              placeholder="如 PEA-4139"),
    FormField("中文名称", kind="text", seq="1.1",
              placeholder="如 水性羟基聚酯-丙烯酸分散体 PEA-4139"),
    FormField("化学品分类", kind="text", seq="1.1",
              placeholder="如 脂肪酸改性的水性羟基聚酯型丙烯酸乳液"),
    FormField("产品使用建议和使用限制", kind="text", seq="1.2",
              placeholder="如 涂料"),
    FormField("供应商名称", kind="text", seq="1.3"),
    FormField("供应商地址", kind="text", seq="1.3"),
    FormField("电话", kind="text", seq="1.3"),
    FormField("传真", kind="text", seq="1.3"),
]

# S3 成分/组成资料: 产品类型 + 成分表 (CAS/含量)
S3_FIELDS: list[FormField] = [
    FormField("产品类型", kind="select", seq="3.1",
              options=["混合物", "单质/化合物", "未知"], required=True),
    FormField("成分", kind="component", seq="3.2",
              placeholder="化学品名称 | CAS编号 | 含量%（w/w）"),
]

# S9 物理和化学特性: 使用者填基础物性 (模板 9.1~9.21)
S9_FIELDS: list[FormField] = [
    FormField("外观", seq="9.1", placeholder="如 乳白色液体"),
    FormField("嗅觉阈值", seq="9.2", placeholder="如 无数据"),
    FormField("pH值", seq="9.3", placeholder="如 7-9"),
    FormField("离子性", seq="9.4", placeholder="如 非离子"),
    FormField("初沸点", seq="9.5", placeholder="如 约100℃"),
    FormField("闪点", seq="9.6", placeholder="如 76℃"),
    FormField("蒸发速率", seq="9.7", placeholder="如 无数据"),
    FormField("可燃性（固态、气态）", seq="9.8", placeholder="如 不适用"),
    FormField("燃烧值", seq="9.9", placeholder="如 不适用"),
    FormField("饱和蒸气压", seq="9.10", placeholder="如 约130hPa在50℃"),
    FormField("相对蒸气密度", seq="9.11", placeholder="如 无数据"),
    FormField("密度", seq="9.12", placeholder="如 约1.06 g/cm3"),
    FormField("水溶性", seq="9.13", placeholder="如 完全混溶"),
    FormField("表面张力", seq="9.14", placeholder="如 无数据"),
    FormField("辛醇/水分配系数对数值", seq="9.15", placeholder="如 无数据"),
    FormField("自燃温度", seq="9.16", placeholder="如 不适用"),
    FormField("引燃温度", seq="9.17", placeholder="如 无数据"),
    FormField("分解温度", seq="9.18", placeholder="如 无数据"),
    FormField("动力粘度", seq="9.19", placeholder="如 ＜400mPa.s"),
    FormField("爆炸特性", seq="9.20", placeholder="如 无数据"),
    FormField("粉尘爆炸级别", seq="9.20", placeholder="如 不适用"),
    FormField("固体含量", seq="9.22", placeholder="如 47±2%（150℃，30min）"),
    FormField("有效成分", seq="9.23", placeholder="如 无数据"),
    FormField("玻璃化温度", seq="9.24", placeholder="如 无数据"),
    FormField("最低成膜温度", seq="9.25", placeholder="如 无数据"),
    FormField("NCO含量", seq="9.26", placeholder="如 无数据"),
    FormField("羟基含量", seq="9.27", placeholder="如 无数据"),
    FormField("熔点/凝固点", seq="9.28", placeholder="如 无数据"),
    FormField("酸值", seq="9.29", placeholder="如 无数据"),
    FormField("碘值", seq="9.30", placeholder="如 无数据"),
    FormField("倾点", seq="9.31", placeholder="如 无数据"),
    FormField("浊点", seq="9.32", placeholder="如 无数据"),
    FormField("分子量", seq="9.33", placeholder="如 无数据"),
    FormField("含量", seq="9.34", placeholder="如 无数据"),
    FormField("APHA值", seq="9.35", placeholder="如 无数据"),
    FormField("HLB值", seq="9.36", placeholder="如 无数据"),
    FormField("其他信息", seq="9.37", placeholder="如 无数据"),
]


def build_form_schema() -> FormSections:
    """返回 PEA-4139 表单输入结构 (S1/S3/S9)."""
    return FormSections(
        s1=S1_FIELDS,
        s3=S3_FIELDS,
        s9=S9_FIELDS,
    )


def derive_s0(s1_values: dict[str, str]) -> list[dict]:
    """从 S1 产品名称/中文名称 派生 S0 页眉页脚字段 (产品名称/产品型号).

    产品型号: 优先用 S1 产品名称; 为空则从中文名称提取
    (如 '水性羟基聚酯-丙烯酸分散体 PEA-4139' → 'PEA-4139').
    Version/公司名称/修订日期 用户交互不填, 保留模板原值.
    """
    import re
    prod = str(s1_values.get("产品名称", "") or "").strip()
    cn = str(s1_values.get("中文名称", "") or "").strip()
    if not prod:
        m = re.search(r"[A-Za-z]{1,6}-?\d{2,}", cn)
        prod = m.group(0) if m else ""
    if not prod:
        return []
    return [{"seq": "", "label": "产品名称", "value": prod},
            {"seq": "", "label": "产品型号", "value": prod}]


def form_to_write_items(schema: FormSections,
                        s1_values: dict[str, str],
                        s3_product_type: str,
                        s3_components: list[ComponentRow],
                        s9_rows: list[tuple[str, str]] | None = None) -> dict:
    """表单值 → 覆写引擎写入项 (仅含使用者填写的 S1/S3/S9).

    s9_rows: S9 动态表格行 [(标签, 值), ...], 序号按表格从上到下自动排 9.1~9.N;
             空标签且空值的整行跳过. 传 None 时回退用 schema.s9 固定字段
             (未填不输出), 用于兼容无表格调用的场景.

    这是表单系统与覆写引擎的对接契约:
      { "sections": {"0": [...], "1": [...], "3": {...}, "9": [...]},
        "keep_structure": [2], "empty_policy": "warn" }
    顶层 keep_structure / empty_policy 与 make_write_items.py 保持一致;
    S0 页眉页脚由 derive_s0 从 S1 产品名称自动派生, 供覆写引擎直接消费;
    推断引擎后续会在此基础上补齐 S2/S4~16。
    """
    def fields(items: list[FormField], values: dict[str, str]) -> list[dict]:
        out = []
        for f in items:
            if f.kind in ("component",):
                continue
            v = str(values.get(f.label, "")).strip()
            if not v:
                v = "无数据"                 # 表单空值统一落为无数据
            out.append({"seq": f.seq, "label": f.label, "value": v})
        return out

    s1 = fields(schema.s1, s1_values)
    s0 = derive_s0(s1_values)

    if s9_rows is None:
        s9 = fields(schema.s9, {})
    else:
        s9 = []
        for i, (label, value) in enumerate(s9_rows, start=1):
            label = str(label or "").strip()
            value = str(value or "").strip()
            if not value:
                if not label:
                    continue                 # 无标签的空白新增行不输出
                value = "无数据"
            s9.append({"seq": f"9.{i}", "label": label, "value": value})

    comps = []
    for c in s3_components:
        if not c.name.strip() and not c.cas.strip() and not c.conc.strip():
            continue
        comps.append({"name": c.name.strip(), "cas": c.cas.strip(), "conc": c.conc.strip()})

    s3 = {"产品类型": s3_product_type.strip(), "components": comps}

    return {"sections": {"0": s0, "1": s1, "3": s3, "9": s9},
            "keep_structure": [2],
            "empty_policy": "warn"}


__all__ = [
    "FormField", "ComponentRow", "FormSections",
    "S1_FIELDS", "S3_FIELDS", "S9_FIELDS",
    "build_form_schema", "form_to_write_items", "derive_s0",
]
