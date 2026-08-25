# -*- coding: utf-8 -*-
"""S11 毒性资料节 — 国标大类归并 (GB/T 16483-2008).

检索工具 (batch_read --matrix) 与入库总表 (build_standard_table) 共用本模块,
保证"检索页面从上到下 → 总表从左到右"完全一致:

- 检索工具把 S11 扩展结构毒理子项 (LD50/物种/方法/NOAEL/组分名等) 归并成
  国标 11 大类别, 不逐个子项平铺成列
- 入库总表 S11 只产 10 大类 + 总结句共 11 列

本模块不依赖 SectionRow / ExtractedField, 只消费文档原始顺序的
(label, value) 序列, 供两类调用方各自适配.
"""
from __future__ import annotations

import re
from typing import Iterable

from core.schema import _clean_label, standard_field_of

# 国标 11 大类别 (GB/T 16483-2008 §11 毒性资料) — 字段名与飞书清单一致
S11_MAJOR_FIELDS = (
    "急性毒性", "主要皮肤刺激性", "主要眼睛刺激性", "致敏性", "致突变性",
    "致癌性", "生殖毒性", "特异性靶器官系统毒性（一次接触/反复接触）",
    "吸入危险", "附加信息",
)
S11_MAJOR_SET = frozenset(S11_MAJOR_FIELDS)
# 大类 → 国标序号 (11.1~11.10): 检索/总表列头统一用 "11.N 大类名"
S11_MAJOR_SEQ = {m: f"11.{i}" for i, m in enumerate(S11_MAJOR_FIELDS, 1)}


# 组分/物质名 (S11/S12 毒理分组): 无值行作为后续子项的**组分前缀标注**.
# 与 NOISE_LABELS 区分: 分组标签 (经口/经皮/体外/体内等) 与毒理残行
# (如 "急性毒性，经皮 聚丙烯酸酯分散体") 只作清洗丢弃, **不**当组分前缀,
# 否则 "急性毒性，经皮 聚丙烯酸酯分散体" 会被错误地拼到后续子项前.
S11_COMPONENT_LABELS = frozenset({
    "羟基聚丙烯酸酯分散体", "乙二醇丁醚", "聚丙烯酸酯分散体", "100 号溶剂油",
    "100号溶剂油", "聚氨酯分散体", "聚醚二胺", "氢化甲醛与苯胺的聚合物",
    "三亚乙基四胺", "2- 丁氧基乙醇", "2-丁氧基乙醇", "号溶剂油",
    "轻质芳香烃石脑油", "丙二醇甲醚", "二乙二醇单丁醚",
})

# 数据来源引导词: 即使 value 为空也**保留**为该大类的说明行 (OS-1330 等
# 扩展结构在 急性毒性经口/经皮/吸入、皮肤/眼睛刺激、遗传毒性 等子组结尾
# 标注 "对类似产品的研究" — 表示该组数据来自对类似产品的研究, 不应丢失).
S11_KEEP_GUIDE_LABELS = frozenset({
    "对类似产品的研究", "对类似产品的毒理学研究", "对类似产品研究",
    "对产品的研究", "对产品的毒理学研究",
})

# 子组标题 → 国标大类 (扩展结构把 STOT 重复接触数据放在 "亚急性，亚慢性和
# 延迟毒性" 子组下; 它不是"附加信息", 归入 特异性靶器官系统毒性).
# 注: 相应地从 core/schema.py S11 附加信息 别名中移除该词, 双轨一致.
_S11_SUBGROUP_MAJOR = {
    "亚急性，亚慢性和延迟毒性": "特异性靶器官系统毒性（一次接触/反复接触）",
}

# 锁定子块标题 → 所属大类 (OS-1330 等扩展结构的 "CMR评估" 子块):
# CMR = 致癌性/致突变性/生殖毒性 综合评估, 其下 "致癌性：基于现有数据，
# 未满足分类标准。" 等行虽以大类名开头, 但只是 CMR 评估的**条目**,
# 不是真正的大类切换. 块内所有子项 (保留 label: value 上下文) 整体归入
# 附加信息, 避免裸句 "基于现有数据，未满足分类标准。" 被拆散到
# 致癌性/致突变性/生殖毒性 各大类而重复出现.
# 注: core/schema.py S11 附加信息 别名同步加入 "CMR评估", 双轨一致.
_S11_BLOCK_MAJOR = {
    "CMR评估": "附加信息",
}


# 噪声标签: 分组行/组分名/泛化接触途径/无信息占位 (S11 归并时作前缀或丢弃,
# 不作为独立大类值).
# 来源: 扩展结构 MSDS (EC-1801/OS-1330/HPU-7651/PA-3110 等) 把组分毒理、
#       GHS 标签要素平铺成独立行 → 这些行是分组/组分名/接触途径.
# 注: 与 tools/build_standard_table.py 旧 _NOISE_LABELS 保持一致 (单一来源).
NOISE_LABELS = frozenset({
    # 分组/泛化标签
    "组分", "产品", "经口", "经皮", "吸入", "食入", "皮肤接触", "眼睛接触",
    "体外", "体内", "毒性", "刺激性", "健康危险信息", "其它危害",
    # 组分/物质名 (S11/S12 毒理分组)
    "羟基聚丙烯酸酯分散体", "乙二醇丁醚", "聚丙烯酸酯分散体", "100 号溶剂油",
    "100号溶剂油", "聚氨酯分散体", "聚醚二胺", "氢化甲醛与苯胺的聚合物",
    "三亚乙基四胺", "2- 丁氧基乙醇", "2-丁氧基乙醇", "号溶剂油",
    "轻质芳香烃石脑油", "丙二醇甲醚", "二乙二醇单丁醚",
    # 毒理标签残行 (值格内混入的通栏残留: "标签+组分名" 残行, 无独立语义)
    "急性毒性，经皮 聚丙烯酸酯分散体", "急性毒性，经口 聚丙烯酸酯分散体",
    # 跨行组合长 label (含句号/空格串联多个概念, 非独立字段)
    "可能造成皮肤过敏反应。 造成严重眼损伤。 吞咽有害。 长期或重复接触可能对器官造成伤害。 可能的接触途径",
    "三亚乙基四胺 对眼睛有严重损害的风险。 兔",
    "三亚乙基四胺 未分类",
    "聚醚二胺 对眼睛有严重损害的风险。 眼睛刺激性试验 兔",
    "聚醚二胺 未分类",
    "氢化甲醛与苯胺的聚合物 无可观察到的效应剂量（NOEL） (大鼠, 经口)",
    "氢化甲醛与苯胺的聚合物 对眼睛有严重损害的风险。 眼睛刺激性试验 兔",
    "根据Buehler（经皮试验）皮肤致敏性",
    # S3 成分表补充引导段 (note 性质)
    "请注意以下物质",
    "特定阈值浓度≥5%",
    "羟基丙烯酸酯聚合物GHS危险性分类",
})


# 纯组分毒理格式 (如 PA-3110) 子项标签 → 国标大类 兜底映射
# 急性毒性: 用 明确词 (LD50/半数致死/纯"毒性"标签/经口·经皮·吸入毒性),
# 不含泛 "毒性" 单字 (否则 "NOAEL（毒性）"/"重复剂量毒性" 会被误判急性毒性).
_S11_SUB_MAP = [
    (re.compile(r"LD\s?50|LC\s?50|半数致死|^急性毒性$|^毒性$|经口毒性|经皮毒性|吸入毒性|皮肤毒性"),
     "急性毒性"),
    (re.compile(r"刺激|腐蚀"), "主要皮肤刺激性"),
    (re.compile(r"致敏|过敏"), "致敏性"),
    (re.compile(r"致突变|Ames|AMES|基因毒|染色体"), "致突变性"),
    (re.compile(r"致癌"), "致癌性"),
    (re.compile(r"生殖|致畸|生育"), "生殖毒性"),
    (re.compile(r"靶器官|STOT"), "特异性靶器官系统毒性"),
    (re.compile(r"吸入危|吞咽"), "吸入危险"),
]


def s11_sub_major(label: str) -> str | None:
    """子项标签 → 国标大类 (兜底, 供无大类上下文的纯组分毒理格式)."""
    for rx, major in _S11_SUB_MAP:
        if rx.search(label):
            return major
    return None


def s11_is_subfield(label: str) -> bool:
    """S11 子项判定: 清洗序号前缀后, 非精确国标大类名一律算子项 (归并不产列).
    别名细分项 (如 "STOT 评估 – 一次性接触" 映射到大类 "特异性靶器官系统毒性")
    也属子项, 只有精确大类名 ("11.1 急性毒性" → "急性毒性") 才产列."""
    return _clean_label(label) not in S11_MAJOR_SET


# 裸占位值 (field value 内 \n 拆行后的通栏残留, 不作独立子项值)
_S11_JUNK = {"无数据资料", "无数据", "专家意见", "不适用", "未提供", "无"}
# 毒理标签残行 (值格内混入的通栏残留, 无独立语义 → 清洗丢弃)
_S11_RESIDUE = frozenset({
    "急性毒性，经皮 聚丙烯酸酯分散体", "急性毒性，经口 聚丙烯酸酯分散体",
    "急性毒性，经皮", "急性毒性，经口", "急性毒性，吸入",
})


def s11_clean_value(v: str) -> str:
    """清洗 field value 内 \n 拆行的裸占位/毒理残行.

    返回拼接后文本; 全为占位/残行 → 返回空串 (由调用方决定丢弃).
    有效数据行 (含组分前缀/标签:值) 原样保留.
    """
    lines = [ln.strip() for ln in (v or "").split("\n")]
    keep = [ln for ln in lines
            if ln and ln not in _S11_JUNK and ln not in _S11_RESIDUE]
    return "\n".join(keep)


# 值语义 → 急性毒性 (补充: OS-1330 等扩展文档把急性毒性结论行
# "评估: 此物质或混合物无急性皮肤毒性" 平铺在生殖毒性子标题后,
# 仅靠 label ("评估") 无法判明大类; 对 value 做急性毒性结论语义识别,
# 避免污染后续大类. 只用明确 "急性*毒性" 结论词 (急性皮肤/呼吸/经口/经皮/吸入毒性),
# 不含 LD50/半数致死 (否则会从生殖毒性区域把 LD50 数据抢走).
_S11_ACUTE_VALUE_RE = re.compile(r"急性(?:皮肤|呼吸|经口|经皮|吸入)?毒性")


def s11_group_rows(rows: Iterable[tuple[str, str]]) -> dict[str, list[str]]:
    """把文档 S11 的字段行 (label, value, 文档原始顺序) 按国标大类归并.

    遇到国标大类标签 (标准字段) → 成为"当前大类", 其后续非大类字段
    (子项) 归入当前大类; 无子项归属 (如"对类似产品研究"等引导残行) 丢弃.
    返回 {大类: [子项值...]}.

    PA-3110 等纯组分毒理格式 (组分名 + LD50/刺激性子项, 无国标大类标签)
    → 用子项标签语义兜底归类: LD50→急性毒性, 刺激性→皮肤刺激,
    组分名作为前缀保留 (避免信息丢失).

    子项归类优先级: 值明确含急性毒性结论 → 急性毒性 (修复 OS-1330
    把 "评估: 此物质或混合物无急性皮肤毒性" 结论行平铺在生殖毒性
    子标题后的污染); 否则归当前大类上下文 (保留生殖/致癌区域的
    LD50/NOAEL 数据不抢走); 无大类上下文 → 用 label 语义兜底
    (LD50→急性毒性, 刺激性→皮肤刺激 等).
    """
    groups: dict[str, list[str]] = {}
    cur: str | None = None
    comp_prefix = ""
    comp_used = False        # 当前组分名前缀是否已标注过 (只标第一条, 避免"无数复制")
    locked_block: str | None = None  # 锁定子块 (CMR评估等): 块内大类名带值行不切换大类
    for label, value in rows:
        if not (label or "").strip():
            continue
        # 子组标题 → 切换大类 (如 "亚急性，亚慢性和延迟毒性" → STOT).
        # 置于 standard_field_of 之前拦截, 覆盖 schema 中残留的错误映射.
        sub_major = _S11_SUBGROUP_MAJOR.get(label)
        if sub_major is not None:
            cur = sub_major
            comp_prefix = ""
            comp_used = False
            locked_block = None
            continue
        # 锁定子块标题 (CMR评估): 进入块 → 归入附加信息, 块内大类名带值行
        # 视为条目 (保留 label: value), 不再触发大类切换.
        block_major = _S11_BLOCK_MAJOR.get(label)
        if block_major is not None:
            cur = block_major
            locked_block = block_major
            comp_prefix = ""
            comp_used = False
            groups.setdefault(cur, []).append(label)   # 保留 "CMR评估" 子块标题
            continue
        f = standard_field_of(11, label)
        if f is not None and f.name in S11_MAJOR_SET:
            if locked_block and (value or "").strip():
                # 锁定块内的 大类名 带值行 (如 CMR 里的 "致癌性：基于现有数据，
                # 未满足分类标准。") 不切换大类, 保留 label: value 上下文作为
                # 块内条目, 避免裸句被拆散到各大类重复出现. 值保留原文
                # (不经过 _S11_JUNK 清洗, 否则 "致癌性: 无数据资料" 会被洗成
                #  "致癌性: " 丢值).
                text = f"{label}: {(value or '').strip()}"
                if comp_prefix and not comp_used:
                    text = comp_prefix + " " + text
                    comp_used = True
                groups.setdefault(cur, []).append(text)
                continue
            is_exact = (label == f.name)
            cur = f.name                       # 切换当前大类
            locked_block = None
            comp_prefix = ""
            comp_used = False
            v = s11_clean_value(value or "")
            if v:
                groups.setdefault(cur, []).append(v)
            elif not is_exact:
                # 大类**子变体** (如 "急性毒性，经口/经皮/吸入"、"原发性皮肤
                # 刺激"、"体外遗传毒性"): 保留原始子组标签行, 体现文档
                # 经口/经皮/吸入 分层 (PEA-4139 等精确大类标签不触发).
                groups.setdefault(cur, []).append(label)
            continue
        # 组分名 (无值, 属 S11_COMPONENT_LABELS 物质名): 作为后续子项前缀
        # (只对组分/物质名生效; 经口/经皮等分组标签与毒理残行不设前缀)
        if not (value or "").strip() and label in S11_COMPONENT_LABELS:
            comp_prefix = label
            comp_used = False
            continue
        v = s11_clean_value(value or "")
        if not v:
            # 数据来源引导词保留 (即使无值): "对类似产品的研究" 等
            if label in S11_KEEP_GUIDE_LABELS and cur is not None:
                groups.setdefault(cur, []).append(label)
            continue
        # 子项归类优先级: 值明确含急性毒性结论 → 急性毒性 (修复 OS-1330
        # 把急性毒性结论行平铺在生殖毒性子标题后导致的污染);
        # 否则归当前大类上下文 (保留生殖/致癌区域的 LD50/NOAEL 数据不抢走);
        # 无大类上下文 (PA-3110 纯组分格式) → 用 label 语义兜底归类.
        if _S11_ACUTE_VALUE_RE.search(v):
            target = "急性毒性"
        elif cur is not None:
            target = cur
        else:
            target = s11_sub_major(label)
        if target is None:
            continue
        # 子项带标签前缀保留语义 (如 "方法: OECD423"), 避免只有半截值
        text = f"{label}: {v}" if label not in NOISE_LABELS else v
        # 组分名前缀只标注该组**第一条**子项, 后续不重复 (OS-1330 等
        # 扩展结构里 "2-丁氧基乙醇" 几乎贯穿每个子组, 逐行加前缀会
        # 无限复制组分名, 用户无法阅读)
        if comp_prefix and not comp_used:
            text = comp_prefix + " " + text
            comp_used = True
        groups.setdefault(target, []).append(text)
    return groups


def s11_value(groups: dict[str, list[str]], label: str) -> str:
    """某国标大类列的取值: 该大类所有子项值按原文顺序合并 (换行)."""
    return "\n".join(groups.get(label, []))
