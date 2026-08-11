# -*- coding: utf-8 -*-
"""MSDS 标准数据库构建器 — 全库 646 份 MSDS → SQLite 三表.

依据 MSDS数据库设计.md 的三表结构:
  1. msds_documents  主表 (一行一文档: S0 页眉 + S1/S4-S8/S10/S12-S14 稳定字段固定列,
                      中英文模板标签统一归一化到标准字段)
  2. msds_components 成分表 (1:N: doc_id + 成分名/CAS/含量)
  3. msds_fields     字段表 (EAV: doc_id + section + field_key + raw_label + value)

S1-S16 全节字段都进 msds_fields (不只 S9/S11), 字段名经 _NORM 归一化
(中英文/异写收敛, 覆盖 S1-S16 标准字段); 主表固定列取每节覆盖≥30% 的稳定字段.

用法:
  python build_db.py "F:\\...\\MSDS" -o msds.db
  python build_db.py "F:\\...\\MSDS" -o msds.db --limit 20   # 只建前 20 个
"""
from __future__ import annotations

import json
import re
import sqlite3
import sys
import time
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))
from core.docx_reader import read_msds
from core.extract import extract_doc, ExtractedField


def _nk(s: str) -> str:
    """归一化键: 去空格/全角/括号/单位等 → 用于模糊匹配."""
    return re.sub(r"[\s_（）()【】\[\]/·℃.%-]", "", (s or "").lower())

_NORM = {
    "外观": ["外观", "外 观", "Appearance", "General Information Appearance: Form",
             "General Information Appearance", "Form", "General Information Form"],
    "气味": ["气味", "Odour", "Odor", "Odour threshold"],
    "嗅觉阀值": ["嗅觉阀值", "Olfactory threshold"],
    "pH值": ["pH值", "pH值（1%水溶液）", "pH值（5%水溶液）", "pH值（10%水溶液）",
             "pH值（1:10稀释在水中）", "pH值（（1:10稀释在水中））",
             "pH value", "pH-value at 20 °C", "pH-value at 25 °C",
             "pH value(1% aqueous solution)", "pH (1:10 diluted with water)",
             "pH value（10% in water）", "pH (10% in water)",
             "pH-value(1% aqueous solution)", "PH value", "PH值", "pH Value", "pH value"],
    "熔点": ["熔点", "Melting point/Melting range",
             "Change in condition Melting point/Melting range",
             "Change in condition Melting point/Melting range: Boiling point/Boiling range"],
    "初沸点": ["初沸点", "最低初沸点", "最低初沸点（℃）", "Initial boiling point",
               "Boiling point/Boiling range", "Initial Boiling Point"],
    "闪点": ["闪点", "闪点（闭口）", "闪点（闭口：）", "闪点（℃）",
             "Flash point", "Flash point (closed)"],
    "蒸发速率": ["蒸发速率", "蒸发速度（醋酸丁酯-1）", "Evaporation rate",
                 "Rate of evaporation", "Evaporation Rate"],
    "可燃性": ["可燃性", "可燃性（固态、气态）", "Flammability (solid, gaseous)",
               "Flammability (Solid, Gaseous)"],
    "爆炸特性": ["爆炸特性", "Explosion characteristics", "Danger of explosion",
                 "Explosion Characteristics"],
    "爆炸极限": ["爆炸极限", "Explosion limits", "Explosion limits: Lower: Upper",
                 "Lower: Upper", "Lower", "Upper"],
    "蒸汽压": ["蒸汽压", "蒸汽压（25℃）", "蒸汽压（20℃，Kpa）",
               "蒸汽压（20℃、mmHg柱）", "Saturated vapor pressure"],
    "密度": ["密度", "Density", "Density at 20 °C",
             "Density at 25°C: Relative density Vapour density"],
    "相对密度": ["相对密度", "Relative density", "比重", "比重/25℃",
                 "Density at 20 °C: Relative density", "Relative Density"],
    "水溶性": ["水溶性", "水中溶解度", "Water solubility",
               "Solubility in / Miscibility with water", "Water Solubility"],
    "表面张力": ["表面张力", "表面张力（10%水溶液）", "Surface tension", "Surface Tension"],
    "辛醇/水分配系数": ["辛醇/水分配系数的对数值", "辛醇/水分配系数对数值",
                       "Logarithm value of octanol/water partition coefficient",
                       "Partition coefficient (n-octanol/water)",
                       "Logarithm of Octanol/Water Partition Coefficient"],
    "自燃温度": ["自燃温度", "Autoignition temperature", "Spontaneous combustion temperature",
                 "Self-igniting", "Self-igniting temperature", "Auto - ignition Temperature"],
    "分解温度": ["分解温度", "Decomposition temperature", "Decomposition Temperature"],
    "引燃温度": ["引燃温度", "Ignition temperature", "Ignition Temperature"],
    "动力粘度": ["动力粘度", "Dynamic viscosity", "Viscosity: Dynamic at 20 °C: Kinematic",
                 "Viscosity: Dynamic at 20 °C", "Dynamic at 20 °C: Kinematic",
                 "Dynamic at 20 °C", "Viscosity: Dynamic at 20 °C: MFFT/℃: Tg/℃",
                 "Viscosity: Dynamic at 20 °C: Hydroxyl value Kinematic",
                 "Viscosity: Dynamic at 20 °C: MFFT℃ Kinematic", "Dynamic Viscosity"],
    "粘度": ["粘度/25℃", "粘度（涂-4杯）", "Viscosity/25℃", "Viscosity", "Kinematic",
             "Kinematic Viscosity"],
    "固体含量": ["固体含量", "Solid content", "Solids content", "Solid Content"],
    "离子性": ["离子性", "Ionic", "Ionicity"],
    "燃烧值": ["燃烧值", "Combustion value", "Heat of combustion", "Combustion Value"],
    "相对蒸气密度": ["相对蒸气密度", "Relative vapor density", "Relative vapour density",
                     "Relative Vapor Density", "Relative Vapour density"],
    "其他信息": ["其他信息", "Other information", "Other Information"],
    "NCO含量": ["NCO含量"],
    "最低成膜温度": ["最低成膜温度MFFT/℃", "MFFT/℃"],
    "玻璃化温度": ["玻璃化温度Tg/℃", "Tg/℃"],
    "有效成分": ["有效成分", "Active Ingredient", "effective constituent", "含量"],
    "百分比挥发性": ["百分比挥发性"],
    "粉尘爆炸级别": ["粉尘爆炸级别", "Dust explosion level", "Dust Explosion Level"],
    "分子量": ["分子量"],
    "颜色": ["Colour"],
    "酸值": ["酸值"],
    "碘值": ["碘值"],
    "羟基含量": ["羟基含量"],
    "倾点": ["倾点"],
    "浊点": ["浊点"],
    "APHA值": ["APHA值"],
    "HLB": ["HLB"],
    "着火点": ["着火点"],
    "饱和蒸气压": ["饱和蒸气压"],
    # ---- S2 危险性概述 (GHS 明细) ----
    "GHS危险性类别": ["GHS危险性类别", "GHS hazard category", "GHS risk category"],
    "物质或混合物分类": ["物质或混合物分类", "GHS分类", "GHS Classification",
                      "Classification of the substance or mixture"],
    "GHS象形图": ["GHS象形图", "GHS-象形图", "GHS pictograms"],
    "信号词": ["信号词", "Signal Word", "Signal word", "警告词", "警示词",
             "警示词 危险", "Warning word", "Warning", "Danger"],
    "危害性说明": ["危害性说明", "危险性说明", "Hazard statements", "Hazardous statements"],
    "防范说明": ["防范说明", "Precautionary statements", "Precautionary measures"],
    "预防措施": ["预防措施", "Prevention", "Precaution and Preventive Measures",
              "Precautions", "Preventive measures"],
    "事故响应": ["事故响应", "Response", "Response to Exposure"],
    "安全储存": ["安全储存", "Safe Storage"],
    "废弃处置": ["废弃处置", "Disposal"],
    "其他危险": ["其他危险", "Other unclassified hazards"],
    "健康危害": ["Health hazards", "Health hazard"],
    "物理化学危害": ["Physical and chemical hazards", "Physical and chemical hazard"],
    "环境危害": ["Environmental hazards", "Environmental hazard"],
    "根据GHS不属于危害化学品": ["根据GHS不属于危害化学品"],
    "标签要素": ["标签要素", "Label Elements", "GHS标签要素"],
    # ---- S3 成分 ----
    "产品类型": ["产品类型", "Product type"],
    "危险组分": ["危险组分", "Dangerous components"],
    "特定阈值浓度": ["特定阈值浓度≥5%"],
    "请注意以下物质": ["请注意以下物质"],
    # ---- S1 标识 (中/英) ----
    "产品名称": ["产品名称", "Product name", "Trade name"],
    "中文名称": ["中文名称"],
    "化学品分类": ["化学品分类", "Chemical category", "Product classification"],
    "产品使用建议": ["产品使用建议和使用限制", "Product use suggestions",
                   "Product use suggestions and restrictions",
                   "Application of the substance / the preparation"],
    "供应商信息": ["供应商信息", "Supplier information", "Further information obtainable from"],
    "供应商名称": ["供应商名称", "Name of supplier", "Manufacturer/Supplier", "Name"],
    "供应商地址": ["供应商地址", "Supplier address", "Head Office Address", "Address"],
    "电话": ["电话", "Tel"],
    "传真": ["传真", "Fax"],
    # ---- S4 急救 (中/英) ----
    "一般措施": ["一般措施", "General measures", "General information"],
    "误服": ["误服", "Mistakenly taken", "After swallowing", "Ingestion", "Misuse"],
    "接触眼睛": ["接触眼睛", "Eye contact", "After eye contact"],
    "接触皮肤": ["接触皮肤", "Contact with skin", "After skin contact"],
    "吸入": ["吸入", "Inhalation", "After inhalation"],
    # ---- S5 消防 (中/英) ----
    "合适的灭火剂": ["合适的灭火剂", "Suitable extinguishing agent", "Suitable extinguishing agents"],
    "不合适的灭火剂": ["不合适的灭火剂", "Unsuitable extinguishing agent", "Unsuitable extinguishing agents"],
    "物质或混合物的特殊危害": ["物质或混合物的特殊危害", "Special hazards of substances or mixtures"],
    "消防预防措施和保护设备": ["消防预防措施和保护设备",
                          "Fire prevention measures and protection equipment",
                          "Protective equipment"],
    # ---- S6 泄漏 (中/英) ----
    "个人预防措施、应急程序": ["个人预防措施、应急程序",
                     "Personal Precautions, Emergency Procedures",
                     "Personal preventive measures and emergency procedures",
                     "Personal preventive and procedures",
                     "Personal Precautionary Measures and Emergency Procedures"],
    "环境保护措施": ["环境保护措施", "Environmental Protection Measures",
                  "Environmental precautions"],
    "污染物收集和清除的方法": ["污染物收集和清除的方法",
                        "Methods for Collecting and Cleaning Up Contaminants",
                        "Methods of pollutant collection and removal",
                        "Methods for Collecting and Removing Contaminants"],
    "泄漏程序": ["Spill and Leak Procedures"],
    "其它节参考": ["Reference to other sections"],
    # ---- S7 操作储存 (中/英) ----
    "安全操作防范": ["安全操作防范", "Handling/Storage Precautions",
                  "Safety operation precautions", "Safe operation precautions"],
    "安全储存条件": ["安全储存条件", "Storage Period and Temperature", "Storage",
                 "Requirements to be met by storerooms and receptacles",
                 "Information about storage in one common storage facility",
                 "Further information about storage conditions",
                 "Safe storage conditions"],
    "防火防爆信息": ["Information about fire - and explosion protection"],
    # ---- S8 接触控制 (中/英) ----
    "呼吸系统防护": ["呼吸系统防护", "Respiratory Protection", "Respiratory protection"],
    "眼睛防护": ["眼睛防护", "Eye Protection", "Eye protection"],
    "身体防护": ["身体防护", "Body protection", "Skin Protection", "Skin protection"],
    "手部防护": ["手部防护", "Hand Protection", "Hand protection", "Protection of hands"],
    "防护手套的合适材料": ["防护手套的合适材料", "Suitable material for protective gloves"],
    "建议": ["建议", "Recommendation", "Additional Protective Measures",
           "Suggestion", "Additional information"],
    "暴露限值": ["暴露限值", "Exposure Limits", "Control parameter/Permissible concentration",
              "Control parameters of components in workplace", "工作场所组分控制参数"],
    "技术设施设计": ["Additional information about design of technical facilities"],
    "一般防护与卫生": ["General protective and hygienic measures"],
    "个人防护设备": ["Personal protective equipment"],
    "需监控组分": ["Ingredients with limit values that require monitoring at the workplace"],
    # ---- S10 稳定性 (中/英) ----
    "化学稳定性": ["化学稳定性", "Chemical stability"],
    "危险分解产物": ["危险分解产物", "Hazardous decomposition products",
                 "Thermal decomposition / conditions to be avoided"],
    "可能的危害反应": ["可能的危害反应", "Possible hazardous reactions"],
    "不相容材料": ["不相容材料", "Incompatible materials"],
    # ---- S12 生态 (中/英) ----
    "生态毒性": ["生态毒性", "Ecotoxicity", "Ecological toxicity"],
    "持久性和降解性": ["持久性和降解性", "Durability and degradability",
                    "Persistence and degradability"],
    "其他": ["其他", "Other", "其他不利的影响"],
    # ---- S13 处置 (中/英) ----
    "处理方法": ["处理方法", "Waste Disposal Method", "Handling method", "Treatment",
              "Waste treatment methods", "Disposal methods", "废弃处理方法"],
    "空容器注意事项": ["空容器注意事项", "Empty Container Precautions",
                    "Uncleaned packaging", "污染包装物", "Contaminated Packaging"],
    "推荐清洁剂": ["Recommended cleansing agents"],
    # ---- S14 运输 (中/英) ----
    "公路和铁路运输": ["公路和铁路运输", "Road and railway transportation", "Road and rail transport"],
    "海上运输": ["海上运输", "Sea transportation", "Sea transport",
              "海上运输： 14.3 空运 14.4 用户特殊注意事项"],
    "空运": ["空运", "Air transportation", "Air Transport",
           "Air transportation: 14.4 Special precautions for users"],
    "用户特殊注意事项": ["用户特殊注意事项", "Special precautions for users", "Special precautions for user"],
    "联合国编号": ["联合国编号", "United Nations number", "UN-Number"],
    "联合国运输名称": ["联合国运输名称", "United Nations shipping name"],
    "运输危险级别": ["运输危险级别", "Transport hazard level"],
    "环境危害": ["环境危害", "Environmental Hazards", "Environmental hazards"],
    "特殊防范措施": ["特殊防范措施", "Special precautions"],
    # ---- S11 毒理学 ----
    "急性毒性": ["急性毒性", "Acute toxicity", "急性毒性，经口", "急性毒性，经皮", "急性毒性，吸入"],
    "主要皮肤刺激性": ["主要皮肤刺激性", "Main skin irritation",
                     "Primary irritant effect on the skin"],
    "主要粘膜刺激性": ["主要粘膜刺激性", "Main mucosal irritation",
                     "Primary irritant effect on the eye"],
    "致敏性": ["致敏性", "Sensitization"],
    "致突变性": ["致突变性", "Mutagenicity"],
    "致癌性": ["致癌性", "Carcinogenicity"],
    "生殖毒性": ["生殖毒性", "Reproductive toxicity", "Reproductive toxicity/生育力"],
    "类似产品的风险评估数据": ["类似产品的风险评估数据", "Risk assessment data of similar products",
                        "Risk assessment data for similar products"],
    "该产品无可用的毒理学研究": ["该产品无可用的毒理学研究。"],
    # ---- S15 法规 (中/英) ----
    "其它的规定": ["其它的规定", "Other regulations", "Other rules", "Other provisions"],
    "符合下列法规要求": ["符合下列法规要求",
                    "Meet the following regulatory requirements",
                    "Comply with the following regulatory requirements"],
    "化学品安全评估": ["Chemical safety assessment"],
    # ---- S16 其他 (英) ----
    "签发部门": ["Department issuing MSDS"],
}

_LOOKUP: dict[str, str] = {}
for _std, _keys in _NORM.items():
    for _k in _keys:
        _LOOKUP[_nk(_k)] = _std

_STD_FIELDS: dict[int, list[tuple[str, str]]] = {
    # (主表列名, 归一化标准字段名)
    1: [("中文名称", "中文名称"), ("化学品分类", "化学品分类"),
        ("产品使用建议", "产品使用建议"), ("供应商名称", "供应商名称"),
        ("供应商地址", "供应商地址"), ("电话", "电话"), ("传真", "传真")],
    4: [("一般措施", "一般措施"), ("误服", "误服"), ("接触眼睛", "接触眼睛"),
        ("接触皮肤", "接触皮肤"), ("吸入", "吸入")],
    5: [("合适灭火剂", "合适的灭火剂"), ("不合适灭火剂", "不合适的灭火剂"),
        ("特殊危害", "物质或混合物的特殊危害"),
        ("消防措施", "消防预防措施和保护设备")],
    6: [("个人预防", "个人预防措施、应急程序"), ("环保措施", "环境保护措施"),
        ("清除方法", "污染物收集和清除的方法"), ("泄漏程序", "泄漏程序")],
    7: [("操作防范", "安全操作防范"), ("储存条件", "安全储存条件")],
    8: [("呼吸防护", "呼吸系统防护"), ("眼睛防护", "眼睛防护"),
        ("身体防护", "身体防护"), ("手部防护", "手部防护"),
        ("手套材料", "防护手套的合适材料"), ("建议", "建议"), ("暴露限值", "暴露限值")],
    10: [("化学稳定性", "化学稳定性"), ("分解产物", "危险分解产物"),
         ("危害反应", "可能的危害反应")],
    12: [("生态毒性", "生态毒性"), ("持久降解", "持久性和降解性"), ("其他", "其他")],
    13: [("处理方法", "处理方法"), ("空容器", "空容器注意事项")],
    14: [("公路铁路", "公路和铁路运输"), ("海上运输", "海上运输"),
         ("空运", "空运"), ("用户注意", "用户特殊注意事项")],
}


def normalize_field(raw_label: str) -> str:
    """原始标签 → 标准字段名 (归一化全节). 未命中返回去前缀原样."""
    lab = (raw_label or "").strip()
    lab = re.sub(r"^[·•\-\s]+", "", lab).strip()      # 去前导项目符号/破折号
    lab = re.sub(r"^\d+(\.\d+)*[\.\s]*", "", lab).strip()    # 去自动编号序号
    return _LOOKUP.get(_nk(lab), lab)


def detect_language(file_name: str) -> str:
    """按文件名判定语言: 含 _EN / -EN / (EN) / (空格)EN 等分隔标记视为英文."""
    return "en" if re.search(r"(?i)(?:^|[_\- (\[])en", file_name) else "zh"


def detect_vendor(path: Path) -> str:
    """按目录结构判定模板厂家 (目录优先): 中文/英文下分 冠志/国彩."""
    p = str(path)
    if "国彩" in p or "guocai" in p.lower() or "guoccai" in p.lower():
        return "国彩"
    if "冠志" in p or "guanzhi" in p.lower():
        return "冠志"
    return "其他"


_SCHEMA = """
CREATE TABLE IF NOT EXISTS msds_documents (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    file_name TEXT NOT NULL,
    file_path TEXT UNIQUE NOT NULL,
    language TEXT,
    template_vendor TEXT,
    product_name TEXT,
    version TEXT,
    revision_date TEXT,
    中文名称 TEXT, 化学品分类 TEXT, 产品使用建议 TEXT,
    供应商名称 TEXT, 供应商地址 TEXT, 电话 TEXT, 传真 TEXT,
    s3_产品类型 TEXT,
    s4_一般措施 TEXT, s4_误服 TEXT, s4_接触眼睛 TEXT, s4_接触皮肤 TEXT, s4_吸入 TEXT,
    s5_合适灭火剂 TEXT, s5_不合适灭火剂 TEXT, s5_特殊危害 TEXT, s5_消防措施 TEXT,
    s6_个人预防 TEXT, s6_环保措施 TEXT, s6_清除方法 TEXT, s6_泄漏程序 TEXT,
    s7_操作防范 TEXT, s7_储存条件 TEXT,
    s8_呼吸防护 TEXT, s8_眼睛防护 TEXT, s8_身体防护 TEXT, s8_手部防护 TEXT,
    s8_手套材料 TEXT, s8_建议 TEXT, s8_暴露限值 TEXT,
    s10_化学稳定性 TEXT, s10_分解产物 TEXT, s10_危害反应 TEXT,
    s12_生态毒性 TEXT, s12_持久降解 TEXT, s12_其他 TEXT,
    s13_处理方法 TEXT, s13_空容器 TEXT,
    s14_公路铁路 TEXT, s14_海上运输 TEXT, s14_空运 TEXT, s14_用户注意 TEXT,
    s15_法规 TEXT, s16_其他信息 TEXT,
    s3_原文 TEXT,
    parse_status TEXT DEFAULT 'ok', parse_error TEXT,
    missing_sections TEXT, anomalies_json TEXT,
    s9_fields_count INTEGER DEFAULT 0,
    s9_fields_json TEXT
);

CREATE TABLE IF NOT EXISTS msds_components (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    doc_id INTEGER NOT NULL REFERENCES msds_documents(id) ON DELETE CASCADE,
    comp_idx INTEGER,
    comp_name TEXT,
    cas TEXT,
    conc TEXT
);
CREATE INDEX IF NOT EXISTS idx_comp_doc ON msds_components(doc_id);

CREATE TABLE IF NOT EXISTS msds_fields (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    doc_id INTEGER NOT NULL REFERENCES msds_documents(id) ON DELETE CASCADE,
    section INTEGER,
    field_key TEXT,
    seq TEXT,
    raw_label TEXT,
    value TEXT
);
CREATE INDEX IF NOT EXISTS idx_field_doc ON msds_fields(doc_id);
CREATE INDEX IF NOT EXISTS idx_field_key ON msds_fields(field_key);
"""


_SIG_BLANK = {"", "Hazard Statements", "Hazard statement", "Precautionary Statements",
              "Precautionary statements", "Signal word", "Signal Word"}


def _fix_signal_word(sec_num: int, raw_label: str, raw_val: str, std: str) -> str:
    """S2 信号词值修正: 标签含值(警示词 危险)或值被误吞(Warning/Danger)."""
    val = raw_val or ""
    if sec_num != 2 or std != "信号词":
        return val
    lab = re.sub(r"^[·•\-\s]+", "", raw_label or "").strip()
    if not val.strip():
        m = re.search(r"(危险|警告|无|Warning|Danger|None)", lab)
        val = m.group(1) if m else val
    elif lab in ("Warning", "Danger") and val.strip() in _SIG_BLANK:
        val = lab
    return val


def _insert_doc(conn, path: Path, r, stats: dict) -> int:
    name = path.name
    lang = detect_language(name)
    vendor = detect_vendor(path)

    fields_by_sec: dict[int, list[ExtractedField]] = {}
    for e in extract_doc(r):
        fields_by_sec.setdefault(e.section, []).append(e)

    header_prod = header_ver = rev_date = ""
    for e in fields_by_sec.get(0, []):
        if "产品名称" in e.label:
            header_prod = e.value or ""
        elif "Version" in e.label:
            header_ver = e.value or ""
        elif "修订日期" in e.label or "Prevision" in e.label or "Revision" in e.label:
            rev_date = e.value or ""

    sec_vals: dict[str, str] = {}
    for sec_num, col_pairs in _STD_FIELDS.items():
        for e in fields_by_sec.get(sec_num, []):
            std = normalize_field(e.label or "")
            for col, stdname in col_pairs:
                if std == stdname:
                    sec_vals[f"s{sec_num}_{col}"] = e.value or ""
                    break

    s15 = s16 = ""
    for e in fields_by_sec.get(15, []):
        if e.value:
            s15 = (s15 + "\n" if s15 else "") + e.value
    for e in fields_by_sec.get(16, []):
        if e.value:
            s16 = (s16 + "\n" if s16 else "") + e.value

    # S3 成分原文: 汇总该节所有 field/note 内容 (成分表格文本化时保底)
    s3_raw_lines = []
    for e in fields_by_sec.get(3, []):
        v = (e.value or "").strip()
        if not v:
            continue
        lab = (e.label or "").strip()
        if e.kind == "note" and not lab:
            s3_raw_lines.append(v)
        elif lab:
            s3_raw_lines.append(f"{lab}: {v}")
    s3_raw = "\n".join(s3_raw_lines)

    row = {
        "file_name": name, "file_path": str(path),
        "language": lang, "template_vendor": vendor,
        "product_name": header_prod, "version": header_ver,
        "revision_date": rev_date,
        "中文名称": sec_vals.get("s1_中文名称", ""),
        "化学品分类": sec_vals.get("s1_化学品分类", ""),
        "产品使用建议": sec_vals.get("s1_产品使用建议和使用限制", ""),
        "供应商名称": sec_vals.get("s1_供应商名称", ""),
        "供应商地址": sec_vals.get("s1_供应商地址", ""),
        "电话": sec_vals.get("s1_电话", ""),
        "传真": sec_vals.get("s1_传真", ""),
        "s3_产品类型": "",
        "s4_一般措施": sec_vals.get("s4_一般措施", ""),
        "s4_误服": sec_vals.get("s4_误服", ""),
        "s4_接触眼睛": sec_vals.get("s4_接触眼睛", ""),
        "s4_接触皮肤": sec_vals.get("s4_接触皮肤", ""),
        "s4_吸入": sec_vals.get("s4_吸入", ""),
        "s5_合适灭火剂": sec_vals.get("s5_合适灭火剂", ""),
        "s5_不合适灭火剂": sec_vals.get("s5_不合适灭火剂", ""),
        "s5_特殊危害": sec_vals.get("s5_特殊危害", ""),
        "s5_消防措施": sec_vals.get("s5_消防措施", ""),
        "s6_个人预防": sec_vals.get("s6_个人预防", ""),
        "s6_环保措施": sec_vals.get("s6_环保措施", ""),
        "s6_清除方法": sec_vals.get("s6_清除方法", ""),
        "s6_泄漏程序": sec_vals.get("s6_泄漏程序", ""),
        "s7_操作防范": sec_vals.get("s7_操作防范", ""),
        "s7_储存条件": sec_vals.get("s7_储存条件", ""),
        "s8_呼吸防护": sec_vals.get("s8_呼吸防护", ""),
        "s8_眼睛防护": sec_vals.get("s8_眼睛防护", ""),
        "s8_身体防护": sec_vals.get("s8_身体防护", ""),
        "s8_手部防护": sec_vals.get("s8_手部防护", ""),
        "s8_手套材料": sec_vals.get("s8_手套材料", ""),
        "s8_建议": sec_vals.get("s8_建议", ""),
        "s8_暴露限值": sec_vals.get("s8_暴露限值", ""),
        "s10_化学稳定性": sec_vals.get("s10_化学稳定性", ""),
        "s10_分解产物": sec_vals.get("s10_分解产物", ""),
        "s10_危害反应": sec_vals.get("s10_危害反应", ""),
        "s12_生态毒性": sec_vals.get("s12_生态毒性", ""),
        "s12_持久降解": sec_vals.get("s12_持久降解", ""),
        "s12_其他": sec_vals.get("s12_其他", ""),
        "s13_处理方法": sec_vals.get("s13_处理方法", ""),
        "s13_空容器": sec_vals.get("s13_空容器", ""),
        "s14_公路铁路": sec_vals.get("s14_公路铁路", ""),
        "s14_海上运输": sec_vals.get("s14_海上运输", ""),
        "s14_空运": sec_vals.get("s14_空运", ""),
        "s14_用户注意": sec_vals.get("s14_用户注意", ""),
        "s15_法规": s15, "s16_其他信息": s16, "s3_原文": s3_raw,
        "missing_sections": json.dumps(
            sorted(n for n in range(1, 17) if n not in r.sections),
            ensure_ascii=False),
        "anomalies_json": json.dumps(
            [{"s": a.section, "lvl": a.level, "msg": a.message}
             for a in r.anomalies], ensure_ascii=False),
    }
    for e in fields_by_sec.get(3, []):
        if normalize_field(e.label or "") == "产品类型":
            row["s3_产品类型"] = e.value or ""
            break

    cols = [k for k, v in row.items() if v is not None]
    ph = ",".join("?" for _ in cols)
    cur = conn.execute(
        f"INSERT INTO msds_documents ({','.join(cols)}) VALUES ({ph})",
        [row[c] for c in cols])
    doc_id = cur.lastrowid

    sec3 = r.sections.get(3)
    comp_added = 0
    if sec3:
        for idx, c in enumerate(sec3.components, 1):
            conn.execute(
                "INSERT INTO msds_components (doc_id, comp_idx, comp_name, cas, conc) "
                "VALUES (?,?,?,?,?)",
                (doc_id, idx, c.name, c.cas, c.conc))
            stats["components"] += 1
            comp_added += 1
    if comp_added == 0:
        # 英文模板文本成分提取: 无标签 note 以项目符号开头的行 = 主要成分
        for e in fields_by_sec.get(3, []):
            if e.kind != "note" or (e.label or "").strip():
                continue
            v = re.sub(r"^[·•\-\s]+", "", (e.value or "")).strip()
            v = re.sub(r"^Description\s*:\s*", "", v, flags=re.I).strip()
            v = re.sub(r"\s*/\s*$", "", v).replace("\xa0", " ")
            if not v:
                continue
            # 过滤"无危险成分"声明 note (OS-8000/OS-8970/PU-2341E/PU-8130
            # 英文模板: 'Hazardous Components' 标题 + 'There are no hazardous
            # components...' 声明 → 非成分, 不能误入成分表)
            low = v.lower()
            if (low in ("void", "hazardous components", "composition")
                    or low.startswith("there are no hazardous components")
                    or low.startswith("no hazardous components")
                    or low.startswith("dangerous components")
                    or low.startswith("composition")):
                continue
            name, conc = v, ""
            # 尝试把尾部含量模式(带 % / w/w / （w/w） / %（w/w）)拆到含量列
            m = re.search(r"\s+(\d+(?:[±＋\-]\d+)?\s*(?:%\s*(?:（w/w）)?|（w/w）|w/w|wt%|W/W))\s*$", v)
            if m:
                conc = m.group(1)
                name = v[:m.start()].rstrip()
            conn.execute(
                "INSERT INTO msds_components (doc_id, comp_idx, comp_name, cas, conc) "
                "VALUES (?,?,?,?,?)",
                (doc_id, comp_added + 1, name, "", conc))
            comp_added += 1
            stats["components"] += 1

    s9_fields: dict[str, str] = {}
    for sec_num in sorted(fields_by_sec):
        if sec_num == 0:
            continue  # S0 页眉页脚已进主表 product_name/version/revision_date
        for e in fields_by_sec[sec_num]:
            if e.kind not in ("field", "note"):
                continue
            std = normalize_field(e.label or "")
            val = _fix_signal_word(sec_num, e.label or "", e.value or "", std)
            if sec_num == 9:
                if std == "其他信息" and not val:
                    continue
                s9_fields[std] = val
            conn.execute(
                "INSERT INTO msds_fields (doc_id, section, field_key, seq, raw_label, value) "
                "VALUES (?,?,?,?,?,?)",
                (doc_id, sec_num, std, e.seq or "", e.label or "", val))
            stats["fields"] += 1

    conn.execute(
        "UPDATE msds_documents SET s9_fields_count=?, s9_fields_json=? WHERE id=?",
        (len(s9_fields), json.dumps(s9_fields, ensure_ascii=False), doc_id))
    return doc_id


def build_db(src_dir: Path, db_path: Path, limit: int = 0) -> dict:
    files = sorted(src_dir.rglob("*.docx"))
    files = [f for f in files if not f.name.startswith("~$")]
    if limit:
        files = files[:limit]
    if db_path.exists():
        db_path.unlink()
    conn = sqlite3.connect(str(db_path))
    conn.executescript(_SCHEMA)
    stats = {"ok": 0, "fail": 0, "docs": 0, "components": 0, "fields": 0}
    for f in files:
        try:
            r = read_msds(f)
        except Exception as exc:
            stats["fail"] += 1
            continue
        _insert_doc(conn, f, r, stats)
        stats["docs"] += 1
        stats["ok"] += 1
    conn.commit()
    conn.close()
    return stats


def main(argv: list[str] | None = None) -> int:
    argv = list(argv if argv is not None else sys.argv[1:])
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    if not argv or "-h" in argv or "--help" in argv:
        print(__doc__)
        return 0
    limit = 0
    out = None
    src = None
    i = 0
    while i < len(argv):
        a = argv[i]
        if a in ("-o", "--out") and i + 1 < len(argv):
            out = argv[i + 1]; i += 2
        elif a == "--limit" and i + 1 < len(argv):
            limit = int(argv[i + 1]); i += 2
        else:
            src = a; i += 1
    if not src or not out:
        print("用法: python build_db.py <MSDS目录> -o 输出.db [--limit N]")
        return 2
    t0 = time.time()
    stats = build_db(Path(src), Path(out), limit)
    print(f"构建完成: {stats['ok']} 文档 / {stats['components']} 成分 / "
          f"{stats['fields']} 字段条目 -> {out} ({time.time()-t0:.1f}s)")
    return 0


if __name__ == "__main__":
    sys.exit(main())