# -*- coding: utf-8 -*-
"""MSDS 标准字段 Schema: 任意写法的 MSDS 统一映射到标准字段视图.

字段集 = 飞书 wiki「5. 父子级结构与字段（17 节完整清单）」权威定义
(https://xcnch7esppuf.feishu.cn/wiki/SfoVwWhPkiQwxvkppyWcrxe4nqo)。

**结构冻结约束 (用户确认 2026-08-17)**:
  - 父子级、标签、宽表列 = 第 5 章清单字段, 清单没有的字段一律不建列;
  - 清单外标签的数据写「无数据」, 不扩充任何结构;
  - aliases 仅用于把不同写法归一化到清单字段, 不引入清单外概念.

reader 产出原始行, Schema 负责归一化/折叠. 供透视总表、标准化覆写输出、
数据库构建直接消费. 不依赖 docx_reader, 可独立调用.
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from .structure import SectionData, SectionRow, split_seq


@dataclass(frozen=True)
class SchemaField:
    """一个标准字段 (每节 Schema 的最小单元)."""
    name: str                 # 标准字段名 (显示/列名, 如 "GHS危险性类别")
    aliases: tuple = ()       # 同义标签 (源文件可能写法, 归一化到 name)
    kind: str = "field"       # field | sub | note | comp
    seq: str = ""             # 标准序号 (如 "2.1"; 空 = 不固定)
    collapse: bool = False    # 空父级分组标签 → 折叠 (不产列, 子级上提)
    multi: bool = False       # 可多行值 (列表字段, 如 防范说明 P 代码)


# 每节标准字段集 = 飞书第 5 章清单字段 (全部 17 节; 清单外字段不收录).
SECTION_SCHEMAS: dict[int, list[SchemaField]] = {
    0: [  # 页眉/页脚
        SchemaField("Version", aliases=("版本",)),
        SchemaField("产品名称"),
        SchemaField("公司名称"),
        SchemaField("产品型号"),
        SchemaField("修订日期"),
        SchemaField("页码"),
    ],
    1: [
        SchemaField("产品名称"),
        SchemaField("中文名称"),
        SchemaField("化学品分类", aliases=("产品类型", "产品分类", "Product type",
                                      "Product classification", "Chemical category")),
        SchemaField("产品使用建议和使用限制", aliases=("使用建议", "用途", "1.2 产品使用建议和使用限制",
                                      "Application of the substance / the preparation",
                                      "Product use suggestions and restrictions")),
        SchemaField("供应商名称", aliases=("Name of supplier", "Manufacturer/Supplier", "Supplier name")),
        SchemaField("供应商地址", aliases=("Supplier address", "Head Office Address")),
        SchemaField("电话", aliases=("Tel", "TEL", "Telephone")),
        SchemaField("传真", aliases=("Fax",)),
    ],
    2: [
        SchemaField("紧急情况概述", kind="sub"),   # 定稿模板 2.1 父级 (参照)
        SchemaField("GHS危险性类别", aliases=("GHS分类", "危险性分类", "GHS危险性", "GHS hazard category",
                                      "CHS hazard category", "Hazard classification", "GHS risk category",
                                      "GHS Classification", "根据GHS不属于危害化学品", "无危险",
                                      "不属于危害化学品", "CHS危险性类别", "GHS 危险性类别")),
        SchemaField("标签要素", aliases=("GHS标签要素", "标签元素", "Labelling according to EU guidelines",
                                    "Label elements"), kind="sub", collapse=True),
        SchemaField("象形图", aliases=("GHS-象形图", "GHS象形图", "危险象形标记", "Pictogram",
                                  "Pictograms", "Hazard pictograms")),
        SchemaField("信号词", aliases=("警示语", "警告词", "Signal word", "Warning word", "警示词")),
        SchemaField("危险性说明", aliases=("危害性说明", "危险说明", "危险声明", "Hazard statements",
                                    "Hazard statement", "Hazard Statement")),
        SchemaField("防范说明", aliases=("Prevention instructions", "Preventive instructions",
                                    "Precautions", "Precautionary statements", "Precautionary measures",
                                    "预防措施", "预防声明", "防范", "事故响应", "泄露", "泄漏", "灭火",
                                    "安全存储", "安全储存", "废弃处置", "入眼睛"), multi=True),
        # 四子列 (2.6~2.9): 「其他危险」拆分, 各自独立父级字段
        SchemaField("物理和化学危险", aliases=("物理化学危险",)),
        SchemaField("健康危害", aliases=("健康危险",)),
        SchemaField("环境危害", aliases=("环境危险",)),
        SchemaField("其他危害", aliases=("其他危险", "其它危害", "其它危险", "没有分类的其他危害")),
    ],
    3: [
        SchemaField("产品类型", aliases=("Mixtures", "单一物质", "Substance", "Product type")),
    ],
    4: [
        SchemaField("一般措施", aliases=("General measures", "General information")),
        SchemaField("误服", aliases=("食入", "After swallowing", "Ingestion", "Swallowed")),
        SchemaField("接触眼睛", aliases=("After eye contact", "Eye contact")),
        SchemaField("接触皮肤", aliases=("After skin contact", "Contact with skin", "Skin contact")),
        SchemaField("吸入", aliases=("After inhalation", "Inhalation")),
    ],
    5: [
        SchemaField("合适的灭火剂", aliases=("灭火剂", "Suitable extinguishing agent", "Suitable extinguishing agents")),
        SchemaField("不合适的灭火剂", aliases=("Unsuitable extinguishing agent", "Unsuitable extinguishing agents")),
        SchemaField("物质或混合物的特殊危害", aliases=("特殊危害", "Special hazards of substances or mixtures")),
        SchemaField("消防预防措施和保护设备", aliases=("防护设备", "Protective equipment",
                                          "Fire prevention measures and protection equipment")),
    ],
    6: [
        SchemaField("个人预防措施、应急程序", aliases=("个人预防措施", "Spill and Leak Procedures",
                                          "Precautions for safe handling")),
        SchemaField("环境保护措施", aliases=("Environmental precautions",)),
        SchemaField("污染物收集和清除的方法", aliases=("收集和清除方法",
                                          "Methods and material for containment and cleaning up")),
    ],
    7: [
        SchemaField("安全操作防范", aliases=("操作注意事项", "安全操作",
                                    "Handling/Storage Precautions", "Information about fire - and explosion protection")),
        SchemaField("安全储存条件", aliases=("储存条件", "Storage", "Storage Period and Temperature",
                                    "Requirements to be met by storerooms and receptacles",
                                    "Information about storage in one common storage facility",
                                    "Further information about storage conditions")),
    ],
    8: [
        SchemaField("暴露控制", kind="sub", collapse=True),
        SchemaField("生物限值", kind="sub", seq="8.2"),   # 子表: 组分名称|标准来源|生物监测指标|生物限值|采样时间
        SchemaField("工程控制", seq="8.3"),
        SchemaField("呼吸系统防护", aliases=("呼吸防护", "Respiratory protection")),
        SchemaField("手部防护", aliases=("Hand Protection", "Protection of hands")),
        SchemaField("防护手套的合适材料", aliases=("手套材料", "Suitable material for protective gloves")),
        SchemaField("氟化橡胶 –FKM", aliases=("氟化橡胶 - FKM", "氟化橡胶-FKM", "氟化橡胶–FKM",
                                    "Fluorinated rubber - FKM", "Fluorinated rubber -FKM")),
        SchemaField("丁基橡胶 –IIR", aliases=("丁基橡胶 - IIR", "丁基橡胶-IIR", "丁基橡胶–IIR",
                                    "Butyl rubber - IIR", "Butyl rubber -IIR")),
        SchemaField("丁腈橡胶 – NBR", aliases=("丁腈橡胶 - NBR", "丁腈橡胶-NBR", "丁腈橡胶–NBR",
                                    "Nitrile rubber - NBR", "Nitrile butadiene rubber - NBR")),
        SchemaField("眼睛防护", aliases=("眼部防护", "Eye protection")),
        SchemaField("身体防护", aliases=("皮肤防护", "Body protection")),
        SchemaField("建议", aliases=("Recommendation", "Suggestion", "Personal protective equipment",
                                "General protective and hygienic measures",
                                "Additional information about design of technical facilities")),
    ],
    9: [
        SchemaField("外观", aliases=("外 观", "Appearance", "Form", "Colour", "Color",
                                "General Information Appearance",
                                "General Information Appearance: Form")),
        SchemaField("嗅觉阈值", aliases=("嗅觉阀值", "气味", "性状", "感官性状",
                                "Odour threshold", "Olfactory threshold", "Odour", "Odor")),
        SchemaField("pH值", aliases=("PH值", "pH", "PH value", "pH-value", "pH Value", "pH value",
                                "pH-value at 20 °C", "pH-value at 25 °C", "pH (10% in water)",
                                "pH value(1% aqueous solution)", "PH-value(1% aqueous solution)",
                                "PH (1:10 diluted with water)", "pH value（10% in water）")),
        SchemaField("离子性", aliases=("Ionicity", "Ionic")),
        SchemaField("初沸点", aliases=("沸点", "最低初沸点", "初沸点/沸点", "Boiling point/Boiling range",
                                "Initial boiling point", "Initial Boiling Point")),
        SchemaField("闪点", aliases=("Flash point", "Flash point (closed)")),
        SchemaField("蒸发速率", aliases=("蒸发速度", "百分比挥发性", "Evaporation rate", "Evaporation Rate",
                                "Rate of evaporation", "Percent volatility", "Volatility",
                                "蒸发速度（醋酸丁酯-1）")),
        SchemaField("可燃性（固态、气态）", aliases=("可燃性", "易燃性", "Flammability (solid, gaseous)",
                                      "Flammability (Solid, Gaseous)")),
        SchemaField("燃烧值", aliases=("Combustion value", "Heat of combustion", "Heat of Combustion")),
        SchemaField("饱和蒸气压", aliases=("蒸气压", "蒸汽压", "Vapour pressure", "Vapor pressure",
                                  "Saturated vapor pressure")),
        SchemaField("相对蒸气密度", aliases=("蒸气密度", "Vapour density", "Relative vapor density",
                                    "Relative Vapor Density", "Relative vapour density")),
        SchemaField("密度", aliases=("Density", "Density at 20 °C", "Density at 25°C",
                                "Relative density (water=1)", "相对密度", "相对密度（水=1）", "比重",
                                "Relative density", "Relative Density",
                                "Density at 20 °C: Relative density")),
        SchemaField("水溶性", aliases=("水中溶解度", "溶解度", "溶解性", "Water solubility", "Water Solubility",
                                "Solubility in / Miscibility with water")),
        SchemaField("表面张力", aliases=("Surface tension", "Surface Tension")),
        SchemaField("辛醇/水分配系数对数值", aliases=("分配系数", "分配系数n-辛醇/水", "辛醇/水分配系数",
                                    "辛醇/水分配系数的对数值", "辛醇/水分配系数对数值",
                                    "辛醇/ 水分配系数的对数值",
                                    "Partition coefficient (n-octanol/water)",
                                    "Logarithm value of octanol/water partition coefficient",
                                    "Logarithm of Octanol/Water Partition Coefficient", "Log Pow")),
        SchemaField("自燃温度", aliases=("自燃点", "着火点", "Ignition temperature",
                                "Autoignition temperature", "Self-igniting",
                                "Spontaneous combustion temperature",
                                "Auto - ignition Temperature", "Self-igniting temperature")),
        SchemaField("引燃温度", aliases=("Ignition Temperature",)),
        SchemaField("分解温度", aliases=("Decomposition temperature",)),
        SchemaField("动力粘度", aliases=("粘度", "Dynamic viscosity", "Viscosity", "Kinematic",
                                "Dynamic at 20 °C", "Dynamic at 20 °C: Kinematic",
                                "Viscosity: Dynamic at 20 °C: Kinematic",
                                "Kinematic Viscosity", "Viscosity: Dynamic at 20 °C",
                                "粘度（涂-4杯）")),
        SchemaField("爆炸特性", aliases=("爆炸极限", "爆炸上下限", "Explosion limits",
                                "Danger of explosion", "Lower", "Upper", "Lower: Upper",
                                "Explosion characteristics")),
        SchemaField("粉尘爆炸级别", aliases=("粉尘爆炸", "Dust explosion level")),
        SchemaField("固体含量", aliases=("Solid content", "Solids content", "Solid Content")),
        # 14 项全量扩充标准字段 (9.23 ~ 9.36):
        SchemaField("有效成分", aliases=("有效成份", "Active Ingredient", "effective constituent",
                                "Active ingredient", " effective constituent")),
        SchemaField("玻璃化温度", aliases=("玻璃化温度Tg/℃", "Tg", "Tg/℃", "玻璃化温度Tg")),
        SchemaField("最低成膜温度", aliases=("最低成膜温度MFFT/℃", "MFFT", "MFFT/℃", "MFFT℃",
                                  "最低成膜温度MFFT")),
        SchemaField("NCO含量", aliases=("异氰酸酯含量", "游离异氰酸酯", "NCO content", "NCO Content")),
        SchemaField("羟基含量", aliases=("羟值", "羟基值", "Hydroxyl value", "Hydroxyl Value")),
        SchemaField("熔点/凝固点", aliases=("熔点", "凝固点", "Melting point/Melting range",
                                    "Melting point", "Change in condition Melting point/Melting range",
                                    "Change in condition")),
        SchemaField("酸值", aliases=("酸价", "Acid value", "Acid Value")),
        SchemaField("碘值", aliases=("碘价", "Iodine value", "Iodine Value")),
        SchemaField("倾点", aliases=("Pour point", "Pour Point")),
        SchemaField("浊点", aliases=("Cloud point", "Cloud Point")),
        SchemaField("分子量", aliases=("Molecular weight", "Molecular Weight")),
        SchemaField("含量", aliases=("纯度", "Content", "Purity")),
        SchemaField("APHA值", aliases=("APHA", "色度", "Hazen")),
        SchemaField("HLB值", aliases=("HLB", "HLB value", "亲水亲油平衡值")),
        # 其他信息必须始终放在 Section 9 序号最末项 (9.37):
        SchemaField("其他信息", aliases=("Other information", "Other Information",
                                "General Information", "General Information Form")),
    ],
    10: [
        SchemaField("化学稳定性", aliases=("稳定性", "Chemical stability")),
        SchemaField("危险分解产物", aliases=("分解产物", "危险的分解产物", "Hazardous decomposition products")),
        SchemaField("可能的危害反应", aliases=("反应性", "危险反应", "Possible hazardous reactions", "Reactivity")),
        SchemaField("应避免的条件", aliases=("避免条件", "Conditions to be avoided",
                                    "Thermal decomposition / conditions to be avoided")),
        SchemaField("禁配物", aliases=("不相容的材料", "不相容物质", "Incompatible materials")),
    ],
    11: [
        SchemaField("急性毒性", aliases=("Acute toxicity", "Acute oral toxicity", "急性毒性，经口", "急性毒性，经皮", "急性毒性，吸入")),
        SchemaField("主要皮肤刺激性", aliases=("皮肤刺激性", "原发性皮肤刺激", "Primary irritant effect", "Primary skin irritation", "on the skin", "皮肤腐蚀/刺激")),
        SchemaField("主要眼睛刺激性", aliases=("眼睛刺激性", "主要粘膜刺激性", "on the eye", "Primary mucosal irritation", "原发性粘膜刺激", "Eye irritation", "严重的眼部损伤/眼部刺激", "严重眼损伤/眼刺激", "严重眼损伤/刺激")),
        SchemaField("致敏性", aliases=("皮肤致敏性", "Sensitization", "Skin sensitization", "皮肤致敏", "呼吸或皮肤过敏")),
        SchemaField("致突变性", aliases=("Mutagenicity", "生殖细胞致突变性", "生殖细胞致突变性物质", "Mutability", "体外遗传毒性", "体内基因毒性")),
        SchemaField("致癌性", aliases=("Carcinogenicity", "IARC专著。致癌性的总体评估")),
        SchemaField("生殖毒性", aliases=("生殖毒性/生育力", "Reproductive toxicity", "生殖毒性（生育能力、未出生小孩）", "Reproductive toxicity/fertility", "生殖毒性/致畸性", "生殖毒性/致畸形", "Teratoxicity", "致畸性")),
        SchemaField("特异性靶器官系统毒性（一次接触/反复接触）", aliases=("特异性靶器官系统毒性", "特定靶器官毒性，单次暴露", "特定靶器官毒性，重复暴露", "特定靶器官毒性-单次曝光", "特定靶器官毒性-重复暴露", "STOT评估-一次性接触", "STOT评估-重复性接触", "STOT 评估 – 一次性接触", "STOT 评估 – 重复性接触", "特异性靶器官毒性-一次接触", "特异性靶器官毒性-反复接触")),
        SchemaField("吸入危险", aliases=("Aspiration hazard", "吸入危害", "吸入性危害", "Inhalation hazard", "吞咽及进入呼吸道可能致命")),
        SchemaField("附加信息", aliases=("Additional toxicological information", "Additional information", "可能接触途径的信息", "慢性影响", "与相关的症状物理、化学和毒理学特征")),
    ],
    12: [
        SchemaField("生态毒性", aliases=("Ecotoxicity", "Aquatic toxicity", "对水生环境有害")),
        SchemaField("持久性和降解性", aliases=("持久性/降解性", "Durability and degradability", "Persistence and degradability")),
        SchemaField("其他不利的影响", aliases=("其他不利影响", "其它不利的影响", "其他", "Other adverse effects", "Additional ecological information", "General notes", "Other", "Others", "通用信息")),
    ],
    13: [
        SchemaField("处理方法", aliases=("废弃处置方法", "Disposal methods", "Handling method",
                                "Recommendation", "Uncleaned packaging", "Recommended cleansing agents", "Void")),
    ],
    14: [
        SchemaField("公路和铁路运输", aliases=("Road and railway transportation",)),
        SchemaField("海上运输", aliases=("Sea transportation",)),
        SchemaField("空运", aliases=("Air transportation",)),
        SchemaField("用户特殊注意事项", aliases=("特殊注意事项", "Special precautions for user",
                                    "Special precautions for users")),
    ],
    15: [
        SchemaField("其它的规定", aliases=("其他规定", "Other rules", "Other regulations", "Other provisions", "Other regulations")),
        SchemaField("符合下列法规要求", aliases=("Comply with the following regulatory requirements",)),
        SchemaField("法规条目", multi=True),   # 任意法规名条目 (EC-1801 等 field 型写法统一)
    ],
    16: [
        SchemaField("免责声明", aliases=("声明", "Disclaimer", "Department issuing MSDS")),
    ],
}

# 基础字段表保留为空 (字段集 = 飞书清单, 无清单外补充字段)
_BASE_SCHEMAS: dict[int, list[SchemaField]] = {}


def standard_fields(num: int) -> list[SchemaField]:
    """该节标准字段列表 (全部节由 SECTION_SCHEMAS 提供)."""
    return SECTION_SCHEMAS.get(num, _BASE_SCHEMAS.get(num, []))


# 同义别名 → 标准字段名 查找表: {节: {别名: 标准名}}
_ALIAS_LOOKUP: dict[int, dict[str, str]] = {}


def _build_alias_lookup() -> None:
    if _ALIAS_LOOKUP:
        return
    for num, flds in list(SECTION_SCHEMAS.items()) + list(_BASE_SCHEMAS.items()):
        m = _ALIAS_LOOKUP.setdefault(num, {})
        for f in flds:
            m[f.name] = f.name
            for a in f.aliases:
                m[a] = f.name


# ------------------------------------------------------------------
# 引导段识别 (S11/S12 等 "以下为XX的参考数据:" 说明段)
# ------------------------------------------------------------------

_GUIDE_RE = re.compile(
    r"^(?!请)(?:以下(?:为|是)?)?[^。；;]{4,}?(?:参考数据|数据|资料|信息|风险评估|评估数据|报告)[：:]\s*$")


def is_guide_line(text: str) -> bool:
    """判断一行是否是引导段 (说明性父级, 非字段标签)."""
    t = (text or "").strip()
    return bool(t) and bool(_GUIDE_RE.match(t))


# ------------------------------------------------------------------
# 归一化 API
# ------------------------------------------------------------------

def _clean_label(raw: str) -> str:
    """去序号前缀 + 去尾冒号 + 压缩空白."""
    _, lbl = split_seq((raw or "").strip())
    return re.sub(r"\s+", " ", lbl).rstrip("：:。 .").strip()


# S9 理化特性 单位括号剥离: 统一字段的关键.
#   "pH值（1%水溶液）" → "pH值"   "闪点（闭口）" → "闪点"
# 剥离后与标准字段 (pH值/闪点/粘度/饱和蒸气压 等) 归并为同一列,
# 避免总库成分列因写法不同而爆炸 (用户核心诉求).
_S9_COND_RE = re.compile(
    r"[\d.]+%|℃|°C|Kpa|kPa|hPa|Pa\b|mPa|g/|kg/|水溶液|固态|气态|闭口|开口|"
    r"\d+:\d+|稀释|在水中|aqueous|diluted|in water")
_S9_UNIT_RE = re.compile(r"[（(]([^（）()]*)[）)]$")   # 尾部括号
_S9_SLASH_RE = re.compile(
    r"[/／]\s*(?:[\d.]+(?:℃|°C)?|[0-9A-Za-z·±/%\-³]{1,12})$")


def _strip_s9_unit(label: str) -> str:
    """剥离 S9 标签尾部通用条件括号/单位, 返回最简形态."""
    t = (label or "").strip()
    # 剥离多层嵌套/双括号, 如 pH值（（1:10稀释在水中））
    t = re.sub(r"[（(]+([^（）()]+)[）)]+$", r"（\1）", t)
    m = _S9_UNIT_RE.search(t)
    if m and _S9_COND_RE.search(m.group(1)):
        t = t[:m.start()].strip()
    t = _S9_SLASH_RE.sub("", t).strip()
    return t


def standard_name(num: int, raw_label: str) -> str:
    """任意原始标签 → 标准字段名 (同义映射; 未命中返回清理后的原名).

    S9 特殊: 直接匹配失败时, 循环剥离单位括号后多次回退匹配.
    """
    _build_alias_lookup()
    lbl = _clean_label(raw_label)
    if not lbl:
        return ""
    lookup = _ALIAS_LOOKUP.get(num, {})
    hit = lookup.get(lbl)
    if hit:
        return hit
    if num == 9:
        t = lbl
        for _ in range(3):
            stripped = _strip_s9_unit(t)
            if stripped == t:
                break
            t = stripped
            hit = lookup.get(t)
            if hit:
                return hit
            for f in SECTION_SCHEMAS.get(9, ()):
                if f.name == t:
                    return t
    return lbl


def standard_field_of(num: int, raw_label: str) -> SchemaField | None:
    """原始标签 → 标准 SchemaField; 未命中返回 None."""
    name = standard_name(num, raw_label)
    for f in standard_fields(num):
        if f.name == name:
            return f
    return None


# ------------------------------------------------------------------
# 折叠 + 标准行序列
# ------------------------------------------------------------------

def iter_standard_rows(sec: SectionData) -> list[SectionRow]:
    """产出 Schema 化后的标准行序列 (折叠空父级, 归一化标签)."""
    rows: list[SectionRow] = []
    last_std: str | None = None
    for row in sec.iter_rows():
        if row.kind == "section":
            rows.append(row)
            last_std = None
            continue
        if row.kind == "field":
            std = standard_name(sec.number, row.label)
            f = standard_field_of(sec.number, row.label)
            if f is not None and f.collapse and not row.value.strip():
                continue
            if std and std == last_std and f is not None and f.multi:
                prev = rows[-1]
                if prev.kind == "field":
                    prev.value = (prev.value + "\n" + row.value).strip()
                continue
            new = SectionRow(kind="field", seq=row.seq, label=std,
                             value=row.value, editable=row.editable,
                             index=row.index, span=row.span)
            rows.append(new)
            last_std = std
            continue
        if row.kind == "sub":
            f = standard_field_of(sec.number, row.label)
            if f is not None and f.collapse:
                last_std = None
                continue
        rows.append(row)
        last_std = None
    return rows


def standard_field_values(sec: SectionData) -> dict[str, list[str]]:
    """该节的标准字段 → 值列表 (同义合并 + 折叠空父级后)."""
    out: dict[str, list[str]] = {}
    for row in iter_standard_rows(sec):
        if row.kind == "field" and row.label:
            out.setdefault(row.label, []).append(row.value)
    return out


def standard_result(result) -> dict[int, dict[str, list[str]]]:
    """整个 ParseResult 的 Schema 化标准字段视图.

    返回 {节号: {标准字段名: [值, ...]}} — 供覆写/建库/透视直接消费.
    """
    return {num: standard_field_values(sec)
            for num, sec in sorted(result.sections.items())}
