# -*- coding: utf-8 -*-
"""Schema 标准字段层测试: 归一化 / 折叠 / 引导段."""
from pathlib import Path

from core.docx_reader import TEMPLATE_PATH, read_msds
from core.schema import (
    is_guide_line,
    iter_standard_rows,
    standard_field_of,
    standard_name,
    standard_result,
)


def test_synonym_mapping_s2():
    """S2 同义写法统一到标准字段 (字段名 = 飞书清单)."""
    assert standard_name(2, "危害性说明") == "危险性说明"
    assert standard_name(2, "GHS-象形图") == "象形图"
    assert standard_name(2, "GHS分类") == "GHS危险性类别"
    # 四子列独立字段
    assert standard_name(2, "物理和化学危险") == "物理和化学危险"
    assert standard_name(2, "其他危险") == "其他危害"


def test_synonym_mapping_s9():
    """S9 理化特性同义归一 (字段名 = 飞书清单, 引燃温度独立列)."""
    assert standard_name(9, "PH值") == "pH值"
    assert standard_name(9, "粘度") == "动力粘度"
    assert standard_name(9, "引燃温度") == "引燃温度"
    assert standard_name(9, "蒸气密度") == "相对蒸气密度"
    assert standard_name(9, "pH值（1%水溶液）") == "pH值"


def test_unknown_label_passthrough():
    """未在 Schema 的标签原样保留 (不丢信息)."""
    assert standard_name(2, "自定义未知字段") == "自定义未知字段"


def test_guide_line_detection():
    """引导段识别 (S11/S12 说明段, 非字段标签)."""
    assert is_guide_line("以下为二乙二醇单丁醚 （CAS号：112-34-5）的生态毒理学参考数据：")
    assert is_guide_line("以下为类似产品的风险评估数据：")
    assert not is_guide_line("生态毒性")
    assert not is_guide_line("急性毒性")


def test_guide_line_not_field_on_read():
    """读取 PEA-4139 模板: 引导段不再是空值 field, 而是跨行 note."""
    p = TEMPLATE_PATH
    if not p.exists():
        return
    r = read_msds(p)
    sec = r.sections.get(12)
    assert sec is not None
    rows = sec.iter_rows()
    guide = [x for x in rows if "参考数据" in x.value]
    assert guide, "应保留引导段文本"
    assert all(x.kind == "note" and x.span for x in guide), \
        "引导段应是跨行 note, 不是 field"


def test_standard_result_merges_synonyms():
    """standard_result 同义字段合并 (危害性说明/危险性说明 → 危险性说明)."""
    p = TEMPLATE_PATH
    if not p.exists():
        return
    r = read_msds(p)
    res = standard_result(r)
    s2 = res.get(2, {})
    # PEA-4139 模板用 危害性说明; 合并后不出现 危险性说明之外的旧名
    assert "危险性说明" in s2
    assert "危害性说明" not in s2
    assert s2["GHS危险性类别"]


def test_standard_parent_labels_s2():
    """S2 正式模板的父级/字段标签应与真实容器结构一致."""
    p = TEMPLATE_PATH
    r = read_msds(p)
    std = [x for x in iter_standard_rows(r.sections.get(2))
           if x.kind != "section"]
    labels = [x.label for x in std]
    assert "物质或混合物分类" not in labels
    assert "标签要素" in labels
    assert next(x.value for x in std if x.label == "标签要素") == "无"
    assert "GHS危险性类别" in labels
