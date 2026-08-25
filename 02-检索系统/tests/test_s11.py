# -*- coding: utf-8 -*-
"""S11 毒理归并回归测试 (core.s11).

覆盖 OS-1330 扩展结构的四类修复:
  1. "对类似产品的研究" 等数据来源引导词不再丢弃
  2. 大类子变体 (急性毒性，经口/经皮/吸入、原发性皮肤刺激等) 保留子组行
  3. 组分名 (2-丁氧基乙醇) 前缀只标注第一条, 不再逐行复制
  4. "亚急性，亚慢性和延迟毒性" 归入 特异性靶器官系统毒性 (STOT),
     不再被 schema 错映射为 附加信息
"""
from __future__ import annotations

from core.s11 import s11_group_rows, s11_value


def _rows(*pairs: tuple[str, str]) -> list[tuple[str, str]]:
    return list(pairs)


def test_keep_guide_labels_acute():
    """急性毒性经口/经皮/吸入 三组保留子组行与"对类似产品的研究"."""
    rows = _rows(
        ("急性毒性，经口", ""),
        ("聚丙烯酸酯分散体", ""),
        ("半数致死剂量(LD50) 大鼠", "> 5,000 mg/kg"),
        ("对类似产品的毒理学研究", ""),
        ("急性毒性，经皮", ""),
        ("评估", "此物质或混合物无急性皮肤毒性"),
        ("对类似产品的研究", ""),
        ("急性毒性，吸入", ""),
        ("评估", "此物质或混合物无急性呼吸毒性"),
        ("对类似产品的研究", ""),
    )
    g = s11_group_rows(rows)
    v = s11_value(g, "急性毒性")
    assert "急性毒性，经口" in v, "经口子组行应保留"
    assert "急性毒性，经皮" in v, "经皮子组行应保留"
    assert "急性毒性，吸入" in v, "吸入子组行应保留"
    assert "对类似产品的毒理学研究" in v, "经口组引导词应保留"
    assert v.count("对类似产品的研究") >= 2, "经皮/吸入组引导词应保留"


def test_keep_guide_label_skin_irritation():
    """原发性皮肤刺激 保留子组行与引导词."""
    rows = _rows(
        ("原发性皮肤刺激", ""),
        ("物种", "家兔"),
        ("结果", "轻微刺激"),
        ("分类", "无皮肤刺激"),
        ("方法", "OECD化学品测试指南404"),
        ("对类似产品的研究", ""),
    )
    g = s11_group_rows(rows)
    v = s11_value(g, "主要皮肤刺激性")
    assert "原发性皮肤刺激" in v, "子组标签行应保留"
    assert "对类似产品的研究" in v, "引导词应保留"


def test_component_prefix_only_first_subitem():
    """组分名 2-丁氧基乙醇 只标注该组第一条子项, 不再逐行复制."""
    rows = _rows(
        ("致癌性", ""),
        ("2-丁氧基乙醇", ""),
        ("NOAEL（毒性）", "125ppm"),
        ("物种", "小鼠，雄性/雌性"),
        ("染毒途径", "吸入性"),
        ("NOAEL（毒性）", "125ppm"),
        ("物种", "大鼠，雄性/雌性"),
    )
    g = s11_group_rows(rows)
    v = s11_value(g, "致癌性")
    assert v.count("2-丁氧基乙醇") == 1, \
        f"组分名应只标注第一条, 实际出现 {v.count('2-丁氧基乙醇')} 次: {v!r}"
    assert "2-丁氧基乙醇 NOAEL（毒性）" in v
    assert "\n物种: 小鼠" in v, "第二条起的子项不再带前缀"


def test_subacute_to_stot_not_additional():
    """亚急性，亚慢性和延迟毒性 → 特异性靶器官系统毒性, 不进附加信息."""
    rows = _rows(
        ("亚急性，亚慢性和延迟毒性", ""),
        ("2- 丁氧基乙醇", ""),
        ("NOAEL", "＜69mg/kg"),
        ("染毒途径", "经口"),
        ("物种", "大鼠，雄性/雌性"),
        ("方法", "OECD化学品测试指南408"),
    )
    g = s11_group_rows(rows)
    assert "NOAEL" in s11_value(g, "特异性靶器官系统毒性（一次接触/反复接触）"), "亚急性数据应归入 STOT"
    assert s11_value(g, "附加信息") == "", "附加信息不应再被错误填充"


def test_exact_major_label_not_duplicated():
    """精确国标大类标签 (PEA-4139 范式) 不额外产生子组行."""
    rows = _rows(
        ("急性毒性", "经口： 半数致死剂量（LD50）/大鼠：约3,306 mg/kg。"),
        ("主要皮肤刺激性", "物种：兔子\n分类：不属于皮肤刺激"),
    )
    g = s11_group_rows(rows)
    assert s11_value(g, "急性毒性") == "经口： 半数致死剂量（LD50）/大鼠：约3,306 mg/kg。"
    assert s11_value(g, "主要皮肤刺激性") == "物种：兔子\n分类：不属于皮肤刺激"


def test_cmr_block_not_split_into_major_groups():
    """CMR评估 子块内的 大类名带值行 不触发大类切换:
    - 致癌性/致突变性/生殖毒性 不接收 CMR 的裸句重复
    - CMR 内容整体归入 附加信息, 保留 label: value 上下文
    """
    rows = _rows(
        ("CMR评估", ""),
        ("2-丁氧基乙醇", ""),
        ("致癌性", "基于现有数据，未满足分类标准。"),
        ("致突变性", "基于现有数据，未满足分类标准。"),
        ("致畸性", "基于现有数据，未满足分类标准。"),
        ("生殖毒性/生育力", "基于现有数据，未满足分类标准。"),
        ("致癌性", "无数据资料"),
        ("致突变性", "基于现有数据，未满足分类标准。"),
    )
    g = s11_group_rows(rows)
    # CMR 内容整体归入附加信息, 保留 label 上下文
    add = s11_value(g, "附加信息")
    assert add.count("CMR评估") == 1
    assert "致癌性: 基于现有数据，未满足分类标准。" in add, "CMR 条目应保留 label 上下文"
    assert "致突变性: 基于现有数据，未满足分类标准。" in add
    assert "致癌性: 无数据资料" in add
    # 各大类不接收 CMR 的裸句重复
    assert s11_value(g, "致癌性") == "", "CMR 的致癌性条目不应进致癌性大类"
    assert s11_value(g, "致突变性") == "", "CMR 的致突变性条目不应进致突变性大类"
    assert s11_value(g, "生殖毒性") == "", "CMR 的生殖条目不应进生殖毒性大类"


def test_cmr_block_component_prefix_once():
    """CMR 块内 2-丁氧基乙醇 只标注第一条, 后续不重复."""
    rows = _rows(
        ("CMR评估", ""),
        ("2-丁氧基乙醇", ""),
        ("致癌性", "基于现有数据，未满足分类标准。"),
        ("致突变性", "基于现有数据，未满足分类标准。"),
        ("致畸性", "基于现有数据，未满足分类标准。"),
        ("生殖毒性/生育力", "基于现有数据，未满足分类标准。"),
    )
    g = s11_group_rows(rows)
    add = s11_value(g, "附加信息")
    assert add.count("2-丁氧基乙醇") == 1, f"组分名应只标一次: {add!r}"


def test_cmr_block_after_major_groups():
    """CMR 块之前的正常大类 (如 STOT/吸入危害) 不被锁定逻辑破坏."""
    rows = _rows(
        ("特异性靶器官系统毒性（一次接触/反复接触）", ""),
        ("2-丁氧基乙醇", "基于现有数据，未满足分类标准。"),
        ("CMR评估", ""),
        ("2-丁氧基乙醇", ""),
        ("致癌性", "基于现有数据，未满足分类标准。"),
        ("致突变性", "基于现有数据，未满足分类标准。"),
    )
    g = s11_group_rows(rows)
    # 锁定块外的 STOT 带值行正常归入 STOT
    assert "基于现有数据，未满足分类标准。" in s11_value(g, "特异性靶器官系统毒性（一次接触/反复接触）")
    # CMR 内容归入附加信息
    assert "致癌性: 基于现有数据" in s11_value(g, "附加信息")
