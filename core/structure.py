# -*- coding: utf-8 -*-
"""MSDS 结构化数据模型.

以 GB/T 16483 16 节为标准, 每个 section 由一个表格承载.
本模块定义读取器输出的统一数据模型, 供显示/比对/覆写指向分析共用.
0 节为页眉页脚, 字段化以便设置 可编辑/不可编辑.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any


@dataclass
class FieldData:
    """一个字段行: 标签 | 值."""
    label: str          # 字段标签 (含编号, 如 "1.1产品名称" / "中文名称")
    value: str          # 字段值 (可为空字符串 = 模板有字段但未填)
    row: int = 0        # 表格内行号 (0 起)
    merged: bool = False  # 是否为跨列合并单元格行
    editable: bool = True  # 可编辑 (可被后续覆写替换); 不可编辑 = 固定字段


@dataclass
class ComponentData:
    """S3 成分表的一行 (已归一化: 全角符号转半角, 多成分已按 \\n 拆分).

    raw_* 保留归一化前的原文 (双轨: 检索/对比用归一化值, 审计/回显用原文).
    - raw_name: 原文名称 (含全角符号/连字符差异)
    - raw_cas:  原文 CAS (含全角横线/占位符原文)
    - raw_conc: 原文含量 (含全角 ＞＜～％)
    """
    name: str = ""
    cas: str = ""
    conc: str = ""
    raw_name: str = ""
    raw_cas: str = ""
    raw_conc: str = ""
    editable: bool = True  # 成分行是否可编辑 (可覆写)


# ============================================================
# S3 成分兼容性归一化层
#
# 全库 644 份 MSDS 的 S3 成分表存在三类格式问题:
#   ① 一行多成分: 名称/CAS/含量三列各自含 \n, 一个三元组实际塞 2-6 个成分
#      (占比约 48% — 覆写比对按成分行对齐时必须先拆分)
#   ② 错用全角符号: 含量 ＞＜～％, 名称全角连字符 －, CAS 错用长横 ——
#   ③ 同一物质写法不统一: N,N-二甲基乙醇胺 vs N,N二甲基乙醇胺 (缺连字符)
# 本层把原始三元组拆分成多个 ComponentData, 并做无损归一化 (保留语义).
# ============================================================

# 全角 → 半角 映射 (成分专用; ℃ × 等保留)
_FW2HW = str.maketrans({
    "＞": ">", "＜": "<", "～": "~", "％": "%", "－": "-",
    "（": "(", "）": ")", "：": ":", "；": ";", "，": ",",
    "　": " ", "＝": "=", "．": ".", "、": ",", "﹪": "%",
})
# 占位/未知 CAS (保留语义, 不当作格式错误)
_CAS_PLACEHOLDERS = {"商业机密", "待确认", "无单一", "无", "不适用", "无数据"}
# 英文占位符 (去空格后形态, 与 _CAS_PLACEHOLDERS 同语义):
# "Trade secret" → "Tradesecret" / "Business secret" → "Businesssecret",
# 原样保留不丢词
_CAS_PLACEHOLDER_EN = {"tradesecret", "businesssecret", "tradesecrets",
                       "businesssecrets", "proprietary",
                       "confidential", "unknown"}
# 含空格的英文占位符 (去空格前匹配, 保留原文本含空格)
# 复数兼容: "Trade Secrets" / "Business Secrets" (PU-3210 英文模板)
_CAS_PLACEHOLDER_EN_SPACE_RE = re.compile(
    r"^(trade\s*secrets?|business\s*secrets?|proprietary|confidential|unknown)$",
    re.I)
# 合法 CAS 形态: 2-7位-2位-1位
_CAS_RE = re.compile(r"^\d{2,7}-\d{2}-\d$")
# 纯横线/符号型无效 CAS (如 "——"、"—"、"－")
_INVALID_DASH_RE = re.compile(r"^[\-—–－~·]+$")


def _to_halfwidth(s: str) -> str:
    """全角符号转半角 (中文/化学名中不转义的部分保持不变)."""
    return (s or "").translate(_FW2HW)


def normalize_component_name(name: str) -> str:
    """成分名归一化: 全角→半角, 压缩空格, 折叠连续连字符, 去首尾标点.

    - "N,N－二甲基乙酰胺(DMAC)"   → "N,N-二甲基乙酰胺(DMAC)"
    - "N,N二甲基乙醇胺"           → "N,N二甲基乙醇胺"   (缺连字符, 保留; 清单归一表处理)
    - " 丙烯酸 聚合物 "           → "丙烯酸聚合物"
    - "2-正辛基-4异噻唑啉--3酮"    → "2-正辛基-4异噻唑啉-3酮"
    """
    t = _to_halfwidth(name or "")
    t = re.sub(r"\s+", " ", t).strip()          # 压缩连续空白
    t = re.sub(r"-{2,}", "-", t)                # 折叠连续连字符
    # 只去尾部多余标点; 括号是化学名有效部分 (如 "(DMAC)" / "（BIT）"), 不剥
    t = t.strip("，,。；;.:：·-—–～~")
    return t


def normalize_component_cas(cas: str) -> str:
    """CAS 号归一化: 去空白, 全角→半角.

    - " 7732-18-5 "              → "7732-18-5"
    - "——" / "－" / "—" / "~·"    → ""     (纯符号 CAS 是格式错误, 置空)
    - "商业机密" / "待确认"         → 原样保留 (有效占位语义, 与"无 CAS"区分)
    - "Trade secret" / "Proprietary" → 原样保留 (英文占位符同语义)
    - "55965-84-9"               → "55965-84-9"   (合法 CAS 原样保留)
    - 含非法字符 (如 "/" 连写两 CAS) → 保留原样 (交由人工核查, 不误拆)
    """
    t0 = (cas or "").strip()
    if _CAS_PLACEHOLDER_EN_SPACE_RE.match(t0):
        return t0                # 英文占位符 (含空格) 原样保留, 词不拆散
    t = _to_halfwidth(t0).replace(" ", "")
    if not t:
        return ""
    if t in _CAS_PLACEHOLDERS or t.lower() in _CAS_PLACEHOLDER_EN:
        return t
    if _INVALID_DASH_RE.match(t):
        return ""
    if _CAS_RE.match(t):
        return t
    return t  # 非规范形态 (连写/待确认) 原样保留, 不丢信息


def normalize_component_conc(conc: str) -> str:
    """含量归一化: 去空白, 全角 ＞＜～％ → 半角 > < ~ %.

    - "＞45"        → ">45"
    - "30.0～40.0"  → "30.0~40.0"
    - "> 30"        → ">30"
    - "95±1"        → "95±1"   (± 保留)
    - "无数据"       → "无数据"  (文字保留)
    """
    t = _to_halfwidth(conc or "").replace(" ", "")
    return t


def split_component_cells(name: str, cas: str, conc: str) -> list[ComponentData]:
    """把一个可能含 \\n 的 S3 成分三元组拆成多个 ComponentData.

    拆分规则 (与全库 48% 多成分单元格的实测结构对齐):
      - 名称/CAS/含量三列各自按 \\n 拆行, 按行索引对齐
      - 行数取三者最大值; 不足行以空串补齐 (如 CAS 单行, 则后续成分 CAS 空)
      - 名称空的行丢弃 (多余 CAS/含量行)
      - 每行逐字段归一化 (全角→半角, 去空白)
    返回至少 1 个 ComponentData (name 可能为空 = 单元格仅含量/CAS).
    """
    names = (name or "").split("\n")
    casses = (cas or "").split("\n")
    concs = (conc or "").split("\n")
    n = max(len(names), len(casses), len(concs))
    names = (names + [""] * n)[:n]
    casses = (casses + [""] * n)[:n]
    concs = (concs + [""] * n)[:n]

    out: list[ComponentData] = []
    for i in range(n):
        nm = normalize_component_name(names[i])
        if not nm:
            continue
        # raw 保留原文 (归一化前), 供审计/回显 (Raw/Normalized 双轨)
        out.append(ComponentData(
            name=nm,
            cas=normalize_component_cas(casses[i]),
            conc=normalize_component_conc(concs[i]),
            raw_name=(names[i] or "").strip(),
            raw_cas=(casses[i] or "").strip(),
            raw_conc=(concs[i] or "").strip(),
        ))
    return out or [ComponentData(name=normalize_component_name(name or ""),
                                 raw_name=(name or "").strip())]


# ============================================================
# 单列(整格)成分表解析
#
# 部分英文/中文 MSDS 的 S3 是单列表格: 整个成分表塞进一个单元格,
# 列间用制表符或空格分隔, 行间用换行分隔. 例:
#   PA-4757:  '3.1 Mixtures\nDescription: Polymer\nDangerous components: Void\n'
#             '3.2 Composition\nIngredient\tCAS\tNo. %(w/w)\n'
#             'Acrylic polymer\t9003-01-4 40±1\nDeionized water 7732-18-5\t60±1'
#   HPU-7651: 'Hazardous Components\n...\nOTHER INGREDIENTS\n'
#             'Concentration\t \t Components\t \t CAS-No.\n'
#             '0.5 - 5%\t Dimethylethanolamine\t 108-01-0\n'
#             'This product contains...'
#   OS-8030:  '3.2 Ingredients \nChemical Name CAS No. % (w/w) \n'
#             'Aliphatic Polycarbodiimide Proprietary 35-45\nWater 7732-18-5 55-65'
#   PA-4408:  '产品类型：\t混合物\n成分\n化学品名称\tCAS 编号\t %（w/w）\n'
#             '羟基丙烯酸酯聚合物\t商业机密\t＞40\n水\t7732-18-5\t＞40\n...'
#
# 识别表头行 (含 CAS 列 + 名称/含量列) 后, 表头前为说明行, 表头后按
# 成分行解析 (抽 CAS → 抽含量 → 余为名称), 无法解析的行为说明行保留.
# ============================================================

# 成分表列语义关键词
_COL_NAME_RE = re.compile(
    r"chemical\s*name|ingredient|component|化学品名称|成分名称|组分|名称|物质", re.I)
_COL_CAS_RE = re.compile(r"cas\s*(?:编号|号|no\.?|number|#)?(?![a-z])", re.I)
_COL_CONC_RE = re.compile(r"concentration|含量|浓度|percent|weight|\bw/w|%", re.I)
# 占位/未知 CAS (保留语义)
_FLAT_CAS_PLACEHOLDER_RE = re.compile(
    r"(proprietary|trade\s*secrets?|商业机密|待确认|无单一|confidential)", re.I)
# 合法 CAS 形态 (在文本中)
_FLAT_CAS_RE = re.compile(r"\b(\d{2,7}-\d{2}-\d)\b")
# 含量形态 (优先级: 范围带% → 单值带% → 带符号单值 → 行尾单值)
_FLAT_CONC_RE = re.compile(
    r"([<>≤≥＞＜]?\s*\d+(?:\.\d+)?\s*[-~±～]\s*\d+(?:\.\d+)?\s*[%％]?|"  # 范围 (35-45 / 0.5 - 5% / 40±1 / 25.0~35.0)
    r"[<>≤≥＞＜]?\s*\d+(?:\.\d+)?\s*[%％]|"                          # 单值带% (>30% / 6.5%)
    r"[<>≤≥＞＜]?\s*\d+(?:\.\d+)?(?=\s|$))"                          # 带符号/行尾单值 (>30 / 6.5)
)


def _flat_col_kind(col: str) -> str:
    """判断成分表表头列语义: 'name' | 'cas' | 'conc' | ''."""
    low = (col or "").strip().lower()
    if _COL_CAS_RE.search(low):
        return "cas"
    if _COL_NAME_RE.search(low) and not _COL_CONC_RE.search(low):
        return "name"
    if _COL_CONC_RE.search(low):
        return "conc"
    return ""


def _flat_extract_cas(text: str) -> tuple[str, str]:
    """从文本抽 CAS (合法编号或占位符), 返回 (cas, 移除后文本).

    英文占位符保留原样含空格 ("Trade secret" 不拆词), 由
    normalize_component_cas 统一保留语义.
    """
    m = _FLAT_CAS_PLACEHOLDER_RE.search(text)
    if m:
        return m.group(1), text.replace(m.group(1), " ", 1)
    m = _FLAT_CAS_RE.search(text)
    if m:
        return m.group(1), text.replace(m.group(1), " ", 1)
    return "", text


def _flat_extract_conc(text: str) -> tuple[str, str]:
    """从文本抽含量, 返回 (conc, 移除后文本)."""
    m = _FLAT_CONC_RE.search(text)
    if m:
        return m.group(1).replace(" ", ""), text.replace(m.group(1), " ", 1)
    return "", text


def _flat_row_to_component(cols: list[str], header_kinds: list[str]) -> tuple | None:
    """把数据行列解析为 (name, cas, conc).

    策略: 表头列语义仅用于判断"该行是否成分行候选"; 实际抽取用全局
    正则 — 拼接各列文本 → 先抽 CAS (合法编号/占位符) → 再抽含量
    → 余为名称. 兼容:
      - 名称/CAS/含量挤一格 (PA-4757: 'Acrylic polymer\t9003-01-4 40±1')
      - 列序为 浓度|名称|CAS (HPU-7651)
      - 空格行 (OS-8030: 'Water 7732-18-5 55-65')
    名称或 (CAS 与含量) 缺失 → None (非成分行, 如尾注/溢出通栏).
    """
    # 表头列映射: 仅判断该行是否含名称/含量列 (成分行候选)
    has_name_col = any(k == "name" for k in header_kinds)
    has_conc_col = any(k in ("conc", "cas") for k in header_kinds)
    if not has_name_col and not has_conc_col:
        return None
    joined = " ".join(cols)
    cas, rest = _flat_extract_cas(joined)
    conc, rest = _flat_extract_conc(rest)
    name = re.sub(r"\s+", " ", rest).strip()
    name = re.sub(r"[\s\-—–]+$", "", name).strip()
    if not name:
        return None
    has_tab = any("\t" in c for c in cols)
    if has_tab:
        # 制表符明确分列: 名称 + (CAS 或含量) 即可
        if not cas and not conc:
            return None
    else:
        # 空格行: 必须抽到 CAS, 否则是说明行 (如 PA-4408 S4 溢出通栏
        # 'GHS 分类：依然液体 3 H226...' 含数字含量特征但无 CAS)
        if not cas:
            return None
    return (name, cas, conc)


def split_flat_component_text(text: str) -> dict | None:
    """把单列(整格)成分表文本拆成结构化 dict; 无法识别表头 → None.

    返回:
      {'pre_lines': [...],   # 表头前说明行 (原始文本)
       'header': [...],      # 表头列
       'header_kinds': [...],  # 每列语义 name/cas/conc
       'rows': [(name, cas, conc), ...],   # 成分行 (原始, 未归一化)
       'post_lines': [...]}  # 表头后非成分说明行
    """
    lines = (text or "").split("\n")
    header_idx = -1
    header_cols: list[str] = []
    header_kinds: list[str] = []
    for i, ln in enumerate(lines):
        t = ln.strip()
        if not t:
            continue
        cols = ([c.strip() for c in t.split("\t") if c.strip()]
                if "\t" in t else t.split())
        if len(cols) < 2:
            continue
        kinds = [_flat_col_kind(c) for c in cols]
        if "cas" in kinds and ("name" in kinds or "conc" in kinds):
            header_idx, header_cols, header_kinds = i, cols, kinds
            break
    if header_idx < 0:
        return None

    pre_lines = [ln.strip() for ln in lines[:header_idx] if ln.strip()]
    rows: list[tuple] = []
    post_lines: list[str] = []
    for ln in lines[header_idx + 1:]:
        t = ln.strip()
        if not t:
            continue
        cols = ([c.strip() for c in t.split("\t") if c.strip()]
                if "\t" in t else t.split())
        comp = _flat_row_to_component(cols, header_kinds)
        if comp is not None:
            rows.append(comp)
        else:
            post_lines.append(t)
    return {"pre_lines": pre_lines, "header": header_cols,
            "header_kinds": header_kinds, "rows": rows,
            "post_lines": post_lines}


def parse_vert_component_table(cells: list[str]) -> dict | None:
    """解析'列内多行成分表': 每列含多段(换行分隔), 每列第1段为表头,
    第2段起为数据行. 例 PU-3210/BEK-500L 英文模板 S3:
      c0 = ['Chemicals', 'Polyurethane dispersion', 'N-ethylpyrrolidone', ...]
      c1 = ['CAS NO.',   'Trade Secrets',            '2687-91-4', ...]
      c2 = ['Content%（w/w）', '25.0~35.0', '0.5~1.5', ...]

    表头须含成分语义列 (name/cas/conc 至少两类); 数据行逐行对齐
    (缺失列补空), 经 _flat_row_to_component 归一化抽 (name, cas, conc).
    返回 {'header': [...], 'rows': [(name, cas, conc), ...]} 或 None.
    """
    if not cells or len(cells) < 2:
        return None
    if not any("\n" in (c or "") for c in cells):
        return None  # 单段行不是列内多行结构
    col_paras: list[list[str]] = []
    for c in cells:
        paras = [p.strip() for p in (c or "").split("\n") if p.strip()]
        if not paras:
            return None
        col_paras.append(paras)
    header = [cp[0] for cp in col_paras]
    header_kinds = [_flat_col_kind(h) for h in header]
    # 表头须含 cas 列 或 (name列 且 conc列) —— 否则不是成分表
    if "cas" not in header_kinds and not ("name" in header_kinds
                                          and "conc" in header_kinds):
        return None
    rows: list[tuple] = []
    max_n = max(len(cp) for cp in col_paras)
    for i in range(1, max_n):
        vals = [cp[i] if i < len(cp) else "" for cp in col_paras]
        # 按列语义直接分派 (已分列, 不做全局抽取避免含量残留混入名称)
        cs = cc = ""
        name_parts: list[str] = []
        for val, kind in zip(vals, header_kinds):
            if kind == "cas":
                cs = val
            elif kind == "conc":
                cc = val
            else:
                name_parts.append(val)
        nm = " ".join(p for p in name_parts if p).strip()
        if nm:
            rows.append((nm, cs, cc))
        else:
            # 名称缺失 → 全局抽取兜底 (如仅 cas/conc 两列)
            comp = _flat_row_to_component(vals, header_kinds)
            if comp is not None:
                rows.append(comp)
    if not rows:
        return None
    return {"header": header, "rows": rows}


@dataclass
class SectionRow:
    """统一表格行: 序号 | 标题 | 内容 三列 + 编辑状态 (供 GUI 渲染与标注).

    序号/标题列为固定区 (不可覆写), 内容列可编辑状态由徽章控制.
    """
    kind: str          # "section" 节标题 | "field" 字段 | "sub" 子标题 | "note" 通栏说明
    seq: str = ""      # 序号列 (如 "8.1", "9.23"; 空 = 无序号)
    label: str = ""    # 标题列 (去序号, 如 "暴露控制")
    value: str = ""    # 内容列
    editable: bool = True
    index: int = 0     # 行内序号 (稳定标识, 供手动标注持久化)


_LINE_SUB_RE = re.compile(r"^\d+\.\d+\s*\S+")
# "标题：" 行尾冒号, 内容在下一行 (如 "GHS分类：\n根据GHS不属于危险物")
_HEADER_END_RE = re.compile(r"^(.+?)[：:]\s*$")
# "标题：内容" 行内冒号, 一行两列
_HEADER_INLINE_RE = re.compile(r"^(.+?)[：:]\s*(.*)$")
# tab 分隔键值对 (无冒号): 单列大块模板用 tab 对齐标签与值
#   (RA-15000 '1.2产品使用建议和使用限\t涂料'). 仅在 tab 前像标签
#   (无长数字、长度≤20) 时拆分, 电话号码等值行 ('86-20-82567990 \t\t传真')
#   不误拆. 见 split_text_block 改进 A.
_TAB_KV_RE = re.compile(r"^(.+?)\t+\s*(.+)$")
# 单字冒号残行: 标签跨行断开的尾字 ('产品使用建议和使用限制' 断成
#   '产品使用建议和使用限' + '制：'). 并入前一字段/pending 的 label 尾字.
#   见 split_text_block 改进 B.
_SINGLE_CHAR_COLON_RE = re.compile(r"^([一-鿿])[：:]\s*$")
# 序号: 数字(可带多级.点) + 可选点 + 标题
# 标题首字符须为 汉字/字母 (中文标题、pH 等英文开头);
# "3 / 5" 不匹配 (序号后接 '/', 防拆页码), "GB/T 16483" 不匹配 (不以数字开头)
_SEQ_HEAD_RE = re.compile(r"^(\d+(?:\.\d+)*)\.?\s*([一-鿿A-Za-z].*)$")

# 无冒号短行 → 标题行 判定 (文本格式段落式 MSDS, 如 BL-8085 S11)
#   "急性毒性，经口" / "原发性皮肤刺激" / "致癌性" 是标题 (短、无句号结尾)
#   "无数据资料" / "专家意见" 是内容词, 不当作标题
_HEADING_FORBIDDEN_WORDS = {"无数据资料", "无数据", "专家意见", "不适用", "未提供", "无"}
_HEADING_END_PUNCT = re.compile(r"[。.！？!?]$")

# S2 标签要素无冒号合一的已知标签前缀: 某些模板全节字段无冒号
# (PA-3225/EC-1800/OS-1330 等), 形如 "GHS-象形图 丙二醇甲醚" —
# 标题=GHS-象形图, 内容=丙二醇甲醚 (有害成分). 有冒号的走 _HEADER_INLINE.
# 注意: 纯标签行 "GHS-象形图" 不在此列 (len(t) <= prefix+1 不触发);
# "必须列在标签上的有害成分" 不是这里拆 (它是 GHS-象形图 的值).
_S2_LABEL_PREFIX = "GHS-象形图"

# S2 标签要素子标题词 (2.2 标签要素, 无编号): 作为独立二级标题,
# 其下 象形图/信号词/危险说明/防范说明 等字段归属 (OS-1310 单列模板).
_S2_LABEL_ELEMENT_WORDS = {"标签元素", "标签要素", "没有其他危险"}
# P 代码子项标签连写 (无冒号): "存储存储已锁定。" → 存储 | 存储已锁定。
_S2_P_SUB_RE = re.compile(r"^(存储|处置|事故响应|废弃处置|补充信息)(.+)$")

# S2 危险分类行 → (分类名, 类别) (懒导入, 避免与 detectors 循环依赖)
#   匹配 "皮肤腐蚀/刺激健康危害  1B类" 等多空格对齐行
_class_split = None


def _split_classification_line(t: str):
    global _class_split
    if _class_split is None:
        from .detectors import split_classification_line
        _class_split = split_classification_line
    return _class_split(t)

# GHS 危险/防范说明代码行 (Hxxx / Pxxx, 可含 '+' 组合): S2 标签要素里
# "H302: Harmful if swallowed." / "P303+P361+P353 如皮肤（或头发）沾染：..."
# 是"危害性说明/防范说明"字段的**内容条目**, 不是节内标题.
# 之前被 _HEADER_INLINE/_END 误拆成独立字段标题 (标题列锁定不可覆写),
# 导致 75 种 H/P 标签在 GUI 三列表 + 覆写指向中错误呈现.
# 此正则拦截这类行 → 一律作内容 (并入前一 pending 标题 或 独立 note).
_GHS_CODE_RE = re.compile(r"^\s*[HP]\s*\d{3}(?:\s*[+＋]\s*\.?\s*[HP]\s*\d{3})*")


def _is_heading_line(t: str, bold: bool = False) -> bool:
    """无冒号行是否像标题行 (文本格式段落式 MSDS 识别).

    条件: 不以句号/感叹号/问号结尾、非内容词.
    - 不加粗: 短 (≤15字), 如 "急性毒性，经口" → 标题;
      "科学地研究，而不仅仅是合理地研究。" → 内容 (句号结尾).
    - 加粗: 放宽长度 (≤60字), 但排除完整陈述句 (含逗号/分号/顿号/长数字),
      如 "Waste Disposal Method" (英文标题) → 标题;
      "根据EC指令2006/121/EG,无可用的接触限值信息" (含逗号+长数字) → 内容.
    """
    if t.startswith("[象形图]"):
        return False            # 图片占位符 → 内容 (并入前一字段/pending), 不当标题
    if t in _HEADING_FORBIDDEN_WORDS:
        return False
    if re.match(r"^[A-Za-z]{1,8}[-\s]?\d", t):
        return False            # 型号/标准号 (RA-15000 / EN 374-3) → 内容, 非标题
    if re.match(r"^\d+[A-Za-z]", t):
        return False            # 数字+字母 (2B类 / 1A类) → 内容, 非标题
    if _HEADING_END_PUNCT.search(t):
        return False
    if bold:
        if len(t) <= 15:
            return True
        # 加粗长行: 排除完整陈述句 (分隔符/长数字表明是句子而非标题)
        return not re.search(r"[,，;；、/]|\d{3,}", t) and len(t) <= 60
    return len(t) <= 15


def split_seq(text: str) -> tuple[str, str]:
    """把 '8.1 暴露控制' 拆成 (序号, 标题); 无序号则 ('', 原文).

    - "8.1 暴露控制"       → ("8.1", "暴露控制")
    - "1.1产品名称"        → ("1.1", "产品名称")
    - "9.3 pH值（1%水溶液）" → ("9.3", "pH值（1%水溶液）")
    - "手部防护"           → ("", "手部防护")
    - "3 / 5"             → ("", "3 / 5")   (页码, 不拆)
    """
    t = (text or "").strip()
    m = _SEQ_HEAD_RE.match(t)
    if m:
        return m.group(1), m.group(2).strip()
    return "", t


def split_line(ln: str) -> tuple[str, str, str]:
    """把一行通栏文本拆成 (kind, label, value) —— 单行版 (取拆出的首行)."""
    rows = split_text_block(ln)
    if not rows:
        return "note", "", ""
    r = rows[0]
    return r.kind, r.label, r.value


def _row_field(label: str, value: str) -> SectionRow:
    """构造 field 行: 序号|标题 拆列."""
    seq, lbl = split_seq(label)
    return SectionRow(kind="field", seq=seq, label=lbl, value=value, editable=True)


def split_text_block(ln: str, bold_rows: set[int] | None = None,
                     section: int | None = None) -> list[SectionRow]:
    """把一段 (可含多行 \\n) 通栏文本拆成多个 序号|标题|内容 三列行.

    目的: 序号/加粗标题 与 内容 分列, 序号/标题列受保护不可覆写, 内容列可覆写.

    - 行尾冒号 "GHS分类：" → 标题待定, 其下无冒号行并入该标题的 value
    - 行内冒号 "氟化橡胶 –FKM:厚度≧0.4mm" → 独立 标题|内容 一行
    - 序号子标题 "8.1 暴露控制" (无冒号) → 序号列 + 标题列, 不可编辑
    - 文本格式段落式 MSDS (BL-8085 S11): 无冒号短行 "急性毒性，经口" /
      "原发性皮肤刺激" → 标题行, 其下内容行配对; 但 "根据GHS不属于危险物"
      这类**完整断言句** (PU-1034 S2, 下一行是另一标题) 应作内容而非标题.
      "对产品的研究." / "科学地研究，而不仅仅是合理地研究。" 等续行 →
      并入上一 field 内容
    - bold_rows: 加粗行索引集合 → 这些行即使超长 (英文标题如
      "Waste Disposal Method") 也优先判为标题行 (仍排除完整陈述句)
    - 其余 → 通栏说明 (内容列)

    返回的 SectionRow 不含 index (由 iter_rows 统一赋值, 保证唯一).
    """
    rows: list[SectionRow] = []
    pending_label: str | None = None
    pending_value: list[str] = []
    lines = [t for t in (ln or "").split("\n") if t.strip()]

    def flush() -> None:
        nonlocal pending_label, pending_value
        if pending_label is not None:
            rows.append(_row_field(pending_label, "\n".join(pending_value)))
        pending_label, pending_value = None, []

    for i, raw in enumerate(lines):
        t = raw.strip()
        # S2 危险分类行: "皮肤腐蚀/刺激健康危害  1B类" (多空格对齐) → 拆成
        #  字段 (分类名=类别). 无值的 大类行 (健康危害/环境危害) 保持标题行
        #  (它们下一行是分类行 → 走 _is_heading_line 成为 pending_kind="sub").
        # 仅 S2 启用, 避免其他节 "名称  值" 对齐文本被误拆.
        if section == 2:
            cls = _split_classification_line(t)
            if cls is not None:
                flush()
                lbl, val = cls
                seq, lbl = split_seq(lbl)
                rows.append(SectionRow(kind="field", seq=seq, label=lbl,
                                       value=val, editable=True))
                continue
            # 标签要素: 无编号子标题 (2.2 标签要素) → sub, 其下挂 信号词/
            # 危险说明/防范说明 等字段. 源文件"标签元素"后直接跟象形图占位,
            # 若并入值会成为挂在上一个分类大标题下的 field, 层级错误.
            if t in _S2_LABEL_ELEMENT_WORDS:
                flush()
                rows.append(SectionRow(kind="sub", seq="", label=t,
                                       value="", editable=False))
                continue
            # 分类大类标题 (健康危害/环境危害): 无值短行, 且下一行是分类行
            # ("皮肤腐蚀/刺激健康危害  1B类") → 本行是大类 → sub 标题.
            # (物理危害/OSHA定义的危险 带值 → 走分类行 → field)
            nxt_cls = (_split_classification_line(lines[i + 1])
                       if i + 1 < len(lines) else None)
            if (nxt_cls is not None and not re.search(r"[：:]", t)
                    and len(t.strip()) <= 15):
                flush()
                seq, lbl = split_seq(t.strip())
                rows.append(SectionRow(kind="sub", seq=seq, label=lbl,
                                       value="", editable=False))
                continue
            # P 代码子项标签连写无冒号 (OS-1310): "存储存储已锁定。" /
            # "处置根据当地/地区/...法规处置内容物/容器。" /
            # "补充信息无。" → 拆 标签|值 (存储/处置/事故响应/废弃处置/补充信息).
            m = _S2_P_SUB_RE.match(t)
            if m:
                flush()
                lbl = m.group(1).strip()
                val = m.group(2).strip()
                seq, lbl = split_seq(lbl)
                rows.append(SectionRow(kind="field", seq=seq, label=lbl,
                                       value=val, editable=True))
                continue
        # GHS H/P 代码行 (H302 / P303+P361+P353 ...) → 内容, 不当标题
        # (中英文 S2 标签要素: 危害性说明/防范说明 的条目, 可能含冒号
        #  如 "P303+P361+P353 如皮肤（或头发）沾染：立即脱掉..." 不应被
        #  _HEADER_INLINE/_END 拆成独立字段标题)
        if _GHS_CODE_RE.match(t):
            if pending_label is not None:
                pending_value.append(t)
            elif rows and rows[-1].kind == "field":
                rows[-1].value = (rows[-1].value + "\n" + t).strip()
            else:
                rows.append(SectionRow(kind="note", label="", value=t,
                                       editable=True))
            continue
        # 改进 B: 单字冒号残行 ('制：') → 并入前一 field/pending 的 label 尾字,
        # 恢复被跨行断开的完整标签 (RA-15000 '产品使用建议和使用限' + '制：'
        # → '产品使用建议和使用限制'). 在 _HEADER_END 之前拦截, 避免残行独立成字段.
        m = _SINGLE_CHAR_COLON_RE.match(t)
        if m:
            piece = m.group(1)
            if pending_label is not None:
                pending_label += piece
            elif rows and rows[-1].kind == "field":
                rows[-1].label = rows[-1].label + piece
            # 无上一字段 → 孤立残行, 忽略
            continue
        m = _HEADER_END_RE.match(t)
        if m:                                   # "标题：" → 内容在下一行
            flush()
            pending_label = m.group(1).strip()
            continue
        m = _HEADER_INLINE_RE.match(t)
        if m and m.group(1).strip():            # "标题：内容" → 独立一行
            # 守卫: 冒号前以多位数开头 (值形态, 如电话 '86-20-82567990 \t\t传真：...')
            # 不是标签行 → 不拆, 落入后续 pending/内容 (并入 '电话：' 的值).
            # '9.5 闪点' 序号 '9' 单数字不受影响; '1.1产品名称' 序号 '1' 后是点.
            if not re.match(r"^\d{2,}", m.group(1).strip()):
                flush()
                rows.append(_row_field(m.group(1).strip(), m.group(2).strip()))
                continue
        # 改进 A: tab 分隔键值对 (无冒号) → 拆 标题|值. 仅当 tab 前像标签
        # (无长数字、长度≤20) 时启用, 避免把电话号码等值行 ('86-20-82567990
        # \t\t传真：...') 误拆成 标题=电话号码. 序号 '1.2产品使用建议和使用限'
        # 拆后 seq='1.2', 后续 '制：' 残行会拼回 label.
        m = _TAB_KV_RE.match(t)
        if m and not re.search(r"\d{2,}", m.group(1)) and len(m.group(1).strip()) <= 20:
            left, right = m.group(1).strip(), m.group(2).strip()
            if re.fullmatch(r"\d+(?:\.\d+)*\.?", left):
                # 纯序号 + tab + 标题 ('3.2\tComposition', PA-4851) →
                # 序号+标题 子标题行, 无值 (tab 后是标题文字而非内容).
                seq, lbl = split_seq(left + " " + right)
                rows.append(SectionRow(kind="sub", seq=seq, label=lbl,
                                       value="", editable=False))
            else:
                flush()
                rows.append(_row_field(left, right))
            continue
        # S2 标签要素无冒号合一: 模板缺陷 (PA-3225/EC-1800/OS-1330 等
        # 全节字段无冒号, 如 "GHS-象形图 丙二醇甲醚" / "信号词 危险").
        # 以已知 S2 标签 + 空格开头 → 拆 标题|值 (标题列锁定).
        if t.startswith(_S2_LABEL_PREFIX) and len(t) > len(_S2_LABEL_PREFIX) + 1:
            flush()
            rows.append(_row_field(_S2_LABEL_PREFIX, t[len(_S2_LABEL_PREFIX):].strip()))
            continue
        if _LINE_SUB_RE.match(t):               # 序号子标题 (无冒号) → 独立 sub 行
            flush()
            seq, lbl = split_seq(t)
            rows.append(SectionRow(kind="sub", seq=seq, label=lbl,
                                   value="", editable=False))
            continue
        # 多位数开头续行 (非小节号): '1106 室' 是上一值 (地址) 的跨行续行,
        # 不是标题 (小节号 '8.1' 含点不受影响; '2 小时' 等一位数数值行也不拦).
        # 修复 PA-4408 地址 '...A 区\n1106 室' 把 '1106 室' 误拆成 标题|序号 的错误.
        if re.match(r"^\d{2,}\s*\S", t):
            if pending_label is not None:
                pending_value.append(t)
            elif rows and rows[-1].kind == "field":
                rows[-1].value = (rows[-1].value + "\n" + t).strip()
            else:
                rows.append(SectionRow(kind="note", label="", value=t,
                                       editable=True))
            continue
        # 无冒号短行 → 标题 or 内容?
        # lookahead 前一行: 若前一行是 行尾冒号标题 ("GHS象形图："), 且当前行
        #   无冒号 → 当前行是它的内容 (PU-1034 "根据GHS不属于危险物" 在
        #   "GHS分类：/GHS象形图：" 后), 并入 pending.
        # lookahead 后一行: 若当前行后紧跟 行尾冒号标题 ("GHS象形图："),
        #   当前行也是内容 —— 除非下一行标题以当前行开头/包含当前行
        #   (8085 "致敏性" → "皮肤致敏性（LLNA）：", 父标题关系).
        # 否则 → 当前行是标题 (8085 "急性毒性，经口" 后跟内容行);
        # 加粗行放宽长度限制 (英文标题 "Waste Disposal Method" 等).
        #
        # 加粗标题例外: 模板中节级大标题/字段标题**必然加粗** (如 S2 的
        #   "物质或混合物分类", S8 的 "工作场所组分控制参数"), 其后紧跟
        #   子标题 ("GHS分类：" / "2.1 物质或混合物的分类") 时, 当前行是
        #   **上一级标题而非内容** —— lookahead"下一行是标题 → 当前行是内容"
        #   的规则只适用于**非加粗**行 (PU-1034 "根据GHS不属于危险物").
        #   加粗标题被误降级会显示在内容列 (note) 或拆成空值 field.
        if _is_heading_line(t, bold=(bold_rows is not None and i in bold_rows)):
            is_bold = bold_rows is not None and i in bold_rows
            prev = lines[i - 1].strip() if i > 0 else None
            nxt = lines[i + 1].strip() if i + 1 < len(lines) else None
            prev_is_header = bool(prev and _HEADER_END_RE.match(prev))
            if prev_is_header and not is_bold:
                if pending_label is not None:
                    pending_value.append(t)
                else:
                    rows.append(SectionRow(kind="note", label="", value=t,
                                           editable=True))
                continue
            nxt_m = _HEADER_END_RE.match(nxt) if nxt else None
            if (nxt_m and not is_bold
                    and not (nxt_m.group(1).strip().startswith(t)
                             or t in nxt_m.group(1).strip())):
                if pending_label is not None:
                    pending_value.append(t)
                else:
                    rows.append(SectionRow(kind="note", label="", value=t,
                                           editable=True))
                continue
            flush()                             # 打断上一 pending, 开启新标题
            pending_label = t
            continue
        if pending_label is not None:           # 无冒号内容行 → 并入待定标题
            pending_value.append(t)
            continue
        if rows and rows[-1].kind == "field":   # 续行 (对产品的研究. 等) → 并入上一 field
            prev = rows[-1]
            prev.value = (prev.value + "\n" + t).strip()
            continue
        rows.append(SectionRow(kind="note", label="", value=t, editable=True))  # 通栏说明
    flush()
    return rows


@dataclass
class SectionData:
    """一个 section 的全部内容 (0 = 页眉页脚, 1..16 = 标准节)."""
    number: int                 # 0..16 (0 = 页眉页脚)
    title: str                  # 节标题 (去编号)
    full_title: str             # 完整标题 (含编号, 如 "1.物料及供应商标识")
    fields: list[FieldData] = field(default_factory=list)
    lines: list[str] = field(default_factory=list)     # 非字段的通栏说明行
    components: list[ComponentData] = field(default_factory=list)  # S3 专用
    component_header: str = ""   # 成分表头实际文本 (如 'Chemical Name | CAS Number | %（w/w）'), 供检索/显示
    row_count: int = 0
    is_component_table: bool = False
    # 原始解析顺序: 'field' | 'line' 按文档行序记录 (字段行与通栏行交替时
    # 保持顺序, 如 S8 子标题 8.1 位于材质行之前). 空 = 手动构造, 用兜底顺序.
    order: list[str] = field(default_factory=list)
    # lines 中加粗行的加粗段索引: {line索引: {段索引}} (供标题列归类;
    # 加粗长行如英文标题 "Waste Disposal Method" 放宽为标题)
    line_bold: dict[int, set[int]] = field(default_factory=dict)

    def field_map(self) -> dict[str, str]:
        """label → value 映射 (后出现的覆盖先出现的)."""
        m: dict[str, str] = {}
        for f in self.fields:
            m[f.label] = f.value
        return m

    def iter_rows(self) -> list[SectionRow]:
        """展开为统一行序列, 按文档原始顺序输出.

        序号/加粗标题 与 内容 分列: 标题列 (sub/field 的 label) 受保护,
        徽章只控制内容列可编辑状态. index 全局递增, 供手动标注持久化.

        字段行与通栏行按 order 保持原始顺序 (如 S8 的 8.1 子标题位于
        材质行之前); 连续 lines 合并为一个块再拆, 保留跨行标题配对
        (如 S15 "符合下列法规要求：" 后法规条目并入其内容列).
        """
        rows: list[SectionRow] = [
            SectionRow(kind="section", label=self.full_title, value="",
                       editable=False, index=0),
        ]
        def _block_bold_rows(gidx_offset: int, items: list[str],
                             line_bold: dict[int, set[int]]) -> set[int]:
            """block 内加粗行索引 (把 line 级段索引映射为 block 行索引)."""
            bold_rows: set[int] = set()
            offset = 0
            for j, ln in enumerate(items):
                for rel in line_bold.get(gidx_offset + j, ()):
                    bold_rows.add(offset + rel)
                offset += len([x for x in ln.split("\n") if x.strip()])
            return bold_rows

        if self.order:
            fi = li = 0
            next_index = len(self.fields)
            block_lines: list[str] = []
            block_start = 0

            def flush_block() -> None:
                nonlocal block_lines, next_index, block_start
                if block_lines:
                    bold_rows = _block_bold_rows(block_start, block_lines,
                                                 self.line_bold)
                    for sr in split_text_block("\n".join(block_lines), bold_rows,
                                               section=self.number):
                        sr.index = next_index
                        next_index += 1
                        rows.append(sr)
                    block_lines = []
                    block_start = li

            for kind in self.order:
                if kind == "field":
                    flush_block()
                    f = self.fields[fi]
                    fi += 1
                    seq, lbl = split_seq(f.label)
                    rows.append(SectionRow(kind="field", seq=seq, label=lbl,
                                           value=f.value, editable=f.editable,
                                           index=fi - 1))
                else:  # "line"
                    block_lines.append(self.lines[li])
                    li += 1
            flush_block()
            return rows
        # 兜底: 无 order (手动构造) → fields 后 lines
        next_index = len(self.fields)
        for i, f in enumerate(self.fields):
            seq, lbl = split_seq(f.label)
            rows.append(SectionRow(kind="field", seq=seq, label=lbl, value=f.value,
                                   editable=f.editable, index=i))
        if self.lines:
            block = "\n".join(self.lines)
            bold_rows = _block_bold_rows(0, self.lines, self.line_bold)
            for sr in split_text_block(block, bold_rows, section=self.number):
                sr.index = next_index
                next_index += 1
                rows.append(sr)
        return rows


@dataclass
class ImageData:
    """嵌入文档的一张图片 (GHS 象形图 / S14 运输标签等).

    blob 为原始字节 (png/jpeg), 供 GUI 显示原图 / 导出文件 / 存库.
    section/table/row/cell 记录来源位置, 与字段 value 中的 [象形图] 占位对应.
    """
    blob: bytes
    ext: str = "png"          # png | jpeg (由 image part 扩展名决定)
    section: int = 0          # 所在节 (0 = 页眉/未确定)
    table: int = 0
    row: int = 0
    cell: int = 0
    width: int = 0            # 显示尺寸 (EMU, 源文件绘制大小)
    height: int = 0


@dataclass
class Anomaly:
    """模板/文档异常报告 (读取器不修改原文件, 只报告)."""
    level: str          # "warn" | "error"
    section: int        # 所在节 (0 = 文档级)
    message: str
    detail: str = ""


@dataclass
class ParseResult:
    """一次完整读取的输出."""
    file_path: str = ""
    file_name: str = ""
    sha256: str = ""
    header: str = ""
    footer: str = ""
    sections: dict[int, SectionData] = field(default_factory=dict)
    anomalies: list[Anomaly] = field(default_factory=list)
    images: list[ImageData] = field(default_factory=list)
    tables_count: int = 0
    paragraphs_count: int = 0
    sections_count: int = 0

    def section(self, num: int) -> SectionData | None:
        return self.sections.get(num)

    def has_section(self, num: int) -> bool:
        return num in self.sections

    def summary(self) -> dict[str, Any]:
        """一行式统计摘要."""
        total_fields = sum(len(s.fields) for s in self.sections.values())
        total_components = sum(len(s.components) for s in self.sections.values())
        return {
            "file": self.file_name,
            "sections": len(self.sections),
            "tables": self.tables_count,
            "fields": total_fields,
            "components": total_components,
            "anomalies": len(self.anomalies),
        }
