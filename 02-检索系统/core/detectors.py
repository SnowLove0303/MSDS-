# -*- coding: utf-8 -*-
"""行类型识别器: 判断一个表格行属于 节标题/字段行/成分行/说明行.

策略: 以 "单元格数 + 冒号特征 + 字段特征词" 为主, 样式为辅,
不依赖固定行号, 兼容国彩/冠志多品牌模板变体.
"""
from __future__ import annotations

import re
from typing import Iterable

# 节标题正则: "1.物料及供应商标识" / "3. 成分/组成资料" / "16. 其他信息"
# 兼容节号前单字母前缀: "v1.物料及供应商标识" (冠志模板第1节版本标记, 全库 233 文件),
# "l1.物料及供应商标识" (定稿模板手误). 前导字母会被剥除.
_SECTION_PREFIX_RE = re.compile(r"^[a-zA-Z]")
SECTION_RE = re.compile(r"^\s*(\d{1,2})\s*[.、．]\s*(.+?)\s*$")
# 英文节标题: "SECTION 1: Identification"
EN_SECTION_RE = re.compile(r"^\s*SECTION\s*(\d{1,2})\s*[:.\-]\s*(.+?)\s*$", re.IGNORECASE)

# S3 成分表头标记 (中/英文)
COMPONENT_HEADERS = (
    ("化学品名称", "CAS编号", "含量"),
    ("成分名称", "CAS号", "含量"),
    ("组分", "CAS", "浓度"),
    ("Ingredient", "CAS", "%"),
    ("Component", "CAS", "w/w"),
)
# 字段标签特征词 (冒号结尾的即字段行; 无冒号但有这些词也判字段)
FIELD_KEYWORDS = ("名称", "分类", "措施", "条件", "防护", "信息", "日期", "编号",
                  "沸点", "闪点", "密度", "粘度", "溶解度", "蒸气压", "温度",
                  "途径", "毒性", "数据", "危害", "危险性", "处理", "运输", "法规",
                  "外观", "气味", "pH", "状态", "颜色", "含量", "描述", "建议", "限制")

# 特征字段 → 节兜底 (某些 MSDS 缺节标题行时, 靠字段特征归属)
SECTION_HINT: dict[str, int] = {
    "急救": 4, "误服": 4, "接触眼睛": 4, "接触皮肤": 4, "吸入": 4,
    "灭火": 5, "消防": 5,
    "泄漏": 6, "收集": 6,
    "操作": 7, "储存": 7,
    "接触控制": 8, "防护": 8, "暴露": 8,
    "物理和化学": 9, "外观": 9, "闪点": 9,
    "稳定性": 10, "反应性": 10, "分解产物": 10,
    "毒性": 11, "毒理": 11,
    "生态": 12, "环境": 12,
    "废弃": 13, "处理注意": 13,
    "运输": 14,
    "法规": 15,
    "其他信息": 16,
}


# S2 危险分类行: "分类名 + ≥2连续空格 + 类别" (OS-1310 等单列模板把
#   "皮肤腐蚀/刺激健康危害              1B类" 对齐排在一行, 无冒号).
# 类别词限定 GHS 分类: 未分类 / 1类/1A类/1B类/2类... / 类别 1 等.
# 匹配 → 拆成字段 (分类名=类别); 不匹配的大类行 (健康危害/环境危害) 保持标题.
_CLASS_ROW_RE = re.compile(
    r"^(?P<n>[^\s].*?)\s{2,}"
    r"(?P<c>(?:未分类|类别\s*\d+(?:[A-Z])?|\d+(?:[A-Z]?B?类)))\s*[。.]?\s*$")


def split_classification_line(text: str) -> tuple[str, str] | None:
    """S2 危险分类行 → (分类名, 类别); 非分类行返回 None."""
    m = _CLASS_ROW_RE.match((text or "").strip())
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return None


def normalize_label(text: str) -> str:
    """字段标签归一化: 去空白/全角冒号修正/去尾部冒号."""
    t = re.sub(r"\s+", " ", (text or "").replace("\xa0", " ")).strip()
    t = t.rstrip("：:。 .")
    return t


def is_section_title(text: str) -> tuple[int | None, str | None]:
    """返回 (节号, 节标题) 若为节标题行; 否则 (None, None).

    兼容:
      - 节号前单字母前缀 (v1./l1.): 剥除前导字母后再匹配.
      - 英文 SECTION 长标题: "SECTION 1: Identification of the substance/
        mixture and of the company/undertaking" (81字符, 超过默认60上限).
    """
    t = re.sub(r"\s+", " ", (text or "").replace("\xa0", " ")).strip()
    if not t:
        return None, None
    # 英文 SECTION 标题可更长 (S1 完整标题可达 80+ 字符); 其余限 60
    if EN_SECTION_RE.match(t):
        if len(t) > 100:
            return None, None
        m = EN_SECTION_RE.match(t)
        n = int(m.group(1))
        if 1 <= n <= 16:
            return n, m.group(2).strip()
    if len(t) > 60:
        return None, None
    # 剥除节号前单个字母前缀 (v1. / l1.), 再按中文节标题匹配
    if _SECTION_PREFIX_RE.match(t):
        t = t[1:].strip()
    m = SECTION_RE.match(t)
    if m:
        n = int(m.group(1))
        title = m.group(2).strip()
        # 排除子编号 ("3.1 混合物") 和字段行 ("9.1  外观：")
        #   字段行 group(2) 以数字开头 (9.1 → 1) 或含冒号
        if 1 <= n <= 16:
            if re.search(r"：|:|\d+\.\d", title):
                return None, None
            if re.match(r"^\d", title):
                return None, None
            return n, title
    return None, None


def extract_section_number(text: str) -> int | None:
    n, _ = is_section_title(text)
    return n


def is_component_header_row(cells: list[str]) -> bool:
    """判断一行是否成分表头 (化学品名称/CAS编号/含量%).

    守卫: 真正的表头行是紧凑单行 — 任一单元格含换行 (整格成分表/长文本)
    即不判定为表头, 避免把单列大单元格 (含 ingredient/cas/% 的整格成分表)
    误判为表头而整体跳过丢失数据. 单元格亦限制长度 (表头标题较短).
    """
    if not cells or not cells[0].strip():
        return False
    if any("\n" in c or len(c) > 60 for c in cells):
        return False
    joined = " ".join(cells).lower()
    # 中/英文表头关键词: (名称/成分列, CAS列, 含量列) 三者齐备
    # chemicals? 覆盖 "Chemical/Chemicals/Chemical name" 复数与省略 (PA-4895)
    name_kw = re.search(r"化学品名称|成分名称|组分|ingredient|component|chemicals?|\bchemical\s*names?", joined)
    cas_kw = re.search(r"cas", joined)
    conc_kw = re.search(r"含量|浓度|concentration|%|w/w|weight", joined)
    if name_kw and cas_kw and conc_kw:
        return True
    # 三列表头精确匹配
    if len(cells) >= 3:
        for h in COMPONENT_HEADERS:
            if cells[0].find(h[0]) >= 0 and (cells[1].find(h[1]) >= 0 or "cas" in cells[1].lower()) and (cells[2].find(h[2]) >= 0 or "%" in cells[2] or "w/w" in cells[2].lower()):
                return True
    return False


# 值形态正则: 数字开头 / 含数值单位 / 颜色+物态 / 化学名结尾
# 注意: 编号+标题 (3.1 产品类型 / 9.5 闪点) 不是值, 在 is_value_shape 里先排除.
# 符号集含全角大于/小于号 (成分含量 '＞40') 与半角 '> <'.
_VALUE_NUM_RE = re.compile(r"^\s*[-−+±＞≥≤<>]?\s*\d")
_VALUE_UNIT_RE = re.compile(
    r"\d\s*(%|％|℃|°C|°F|mPa·?s|mPa\s*s|g/?cm3|g/?cm|g/?m3|g/?m|kg/?m3|kg/?m|"
    r"mg/?L|mg/?m3|mg/?kg|ppm|mol/?L|N/?m|S/?cm|Pa\s*s|kJ|MJ|hPa|mm|mbar|bar|µm|μm)")
_COLOR_STATE_RE = re.compile(
    r"^(?:乳白|乳黄|白|黑|红|黄|蓝|绿|灰|棕|浅|深|无|微|淡)?"
    r"[^\s，。]{0,6}(?:液体|乳液|粉末|粉状|固体|气体|气态|透明|不透明|胶体|糊状|膏状|色)$")
# 化学名: 前缀(1-6字/英文) + 结尾词 (乙醇 = 乙+醇; 乙二醇丁醚 = 乙二醇丁+醚)
_CHEM_NAME_RE = re.compile(
    r"^(?:[A-Za-z][A-Za-z0-9\-]*|[一-鿿]{1,6})"
    r"(?:醇|油|酯|烷|酸|胺|酮|醚|烃|苯|烯|醛|酚|橡胶|分散体|乳液|树脂|溶液|溶剂油|蜡)$")
_GLOVE_MAT_RE = re.compile(r"^(?:天然|丁腈|丁基|氟化|氯丁|氯磺化|聚氯乙烯|PVC|NBR|IIR|FKM|CR|NR)[-\s]*橡胶")


def is_value_shape(text: str) -> bool:
    """左栏是否为"值"形态 (数值/百分比/单位/颜色+物态/化学名/手套材质).

    用于区分"加粗的值行"与"字段标签": 值行常被模板加粗, 若不加守卫会被
    bold0 兜底误判为字段标签 (如 '42±2%', '乳白色液体', '乙醇', '丁腈橡胶–NBR').
    """
    t = (text or "").strip()
    if not t or len(t) > 24:
        return False
    if "：" in t or ":" in t:
        return False          # 含冒号必是标签
    if any(kw in t for kw in FIELD_KEYWORDS):
        return False          # 含字段特征词必是标签 (不误拦 '可燃性（固态、气态）')
    # 编号+标题 (3.1 产品类型 / 9.5 闪点 / 8.2 暴露控制) 不是值
    # 排除1: 点后空格+任意 (3.1 产品类型 / 14.1 UN-Number)
    # 排除2: 无空格编号+汉字标题 (1.2产品使用建议 / 9.5闪点 / 3.2成分);
    #        数字后接字母是值 (1.06g/cm3) 不排除
    if re.match(r"^\s*\d+\.\d+\s+\S", t):
        return False
    if re.match(r"^\s*\d{1,2}\s*[.、．]\s*\d{1,2}\s*[一-鿿]", t):
        return False
    # 排除3: 无空格编号+纯英文字母标题 (14.1UN-Number / 2.1Classification).
    #        中间段要求 ≥1 字符: 两字母单位值 (1.5kg) 中间为空不被拦;
    #        含单位斜杠的值 (1.06g/cm3) 因 '/' 不在字符集不被拦.
    if re.match(r"^\s*\d+\.\d+\s*[A-Za-z][A-Za-z0-9\-]+[A-Za-z]\s*$", t):
        return False
    if _VALUE_NUM_RE.match(t):
        return True
    # 非数字开头的单位命中: 仅限短串 ('约50%' / '20 °C'), 长串多为英文标签
    # 含温度/单位条件 ('Density at 20 °C' 是字段标签, 非值).
    if _VALUE_UNIT_RE.search(t) and len(t) <= 12:
        return True
    if _COLOR_STATE_RE.match(t):
        return True
    if _CHEM_NAME_RE.match(t):
        return True
    if _GLOVE_MAT_RE.match(t):
        return True
    # 纯数字/编码 (EN 374-3, UN No)
    if re.match(r"^[A-Z]{1,6}[-\s]?\d", t):
        return True
    return False


def looks_like_field_label(text: str) -> bool:
    """加粗兜底守卫: 左栏是否像字段标签 (而非加粗的值/说明).

    bold0 兜底只应把"确实像标签"的加粗行判为字段; 明显是值形态
    (数值/物态/化学名) 的加粗行应走说明行, 避免 '42±2%', '乳白色液体' 误判.
    """
    t = (text or "").strip()
    if not t:
        return False
    if len(t) > 30:
        return False
    if "：" in t or ":" in t:
        return True
    if any(kw in t for kw in FIELD_KEYWORDS):
        return True
    # 项目符号子行 (如 '· After inhalation') 归入前一字段值, 非独立标签
    if re.match(r"^[·•\-]\s", t):
        return False
    if is_value_shape(t):
        return False
    return True   # 无特征词、非值形态的普通词默认按标签处理 (兼容 'Fire', 'Odor' 等)


def is_field_row(cells: list[str]) -> bool:
    """判断一行是否 字段: 值 对.

    条件: 左栏非空 且 (含冒号 或 含字段特征词).
    说明行 (单列通栏) 返回 False.
    """
    if not cells or not cells[0]:
        return False
    left = cells[0].strip()
    if len(cells) == 1:
        # 单列: 只有含特征词的才勉强算字段 (如 "8.1 暴露控制" 子标题)
        return False
    if "：" in left or ":" in left:
        return True
    # 无冒号但像字段 (如 "8.1 暴露控制")
    low = left.lower()
    for kw in FIELD_KEYWORDS:
        if kw.lower() in low:
            return True
    return False


def is_sub_heading(cells: list[str]) -> bool:
    """判断是否为节内子标题 (无值, 加粗, 形如 '8.1 暴露控制')."""
    if not cells or not cells[0]:
        return False
    left = cells[0].strip()
    if len(cells) > 1 and cells[1].strip():
        return False  # 有值就不是子标题
    # 形如 "8.1" / "2.1" 开头且无冒号
    return bool(re.match(r"^\d+\.\d+\s*\S+", left)) and not re.search(r"：|:", left)


def hint_section(text: str) -> int | None:
    """用特征字段词判断一行属于哪一节 (兜底)."""
    t = text.strip().lower()
    for key, num in SECTION_HINT.items():
        if key.lower() in t:
            return num
    return None


def dedupe_merged(cells: Iterable[str]) -> list[str]:
    """去重合并单元格造成的相邻重复 (横向合并)."""
    out: list[str] = []
    for c in cells:
        c = (c or "").strip()
        if not out or c != out[-1]:
            out.append(c)
    return out
