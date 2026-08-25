# -*- coding: utf-8 -*-
"""MSDS 合格标准判定: 以 PEA-4139 结构为模板, 筛出"完美合格"的标准库候选.

依据 docs/17节合格标准.md. 判定分级:
  - hard:   必需字段 (缺失/空 = 该节不合格)
  - soft:   软性字段 (缺失记 warning, 不计硬伤)
  - veto:   否决项 (出现即该节不合格, 如跨节错位/标签粘连/H·P代码缺失)

API:
  evaluate(result) -> ComplianceReport  单文件判定
  evaluate_dir(in_dir) -> 汇总全库合格/不合格清单
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field as dc_field
from pathlib import Path

from .docx_reader import read_msds
from .schema import SECTION_SCHEMAS, _BASE_SCHEMAS, standard_name


# ------------------------------------------------------------------
# 各节合格标准 (hard 必需字段 / soft 软性字段)
# 字段名 = Schema 标准名 (归一化后匹配)
# ------------------------------------------------------------------

# S2 否决: 跨节内容特征词 (急救/消防/泄漏/储存/废弃 等)
_S2_CROSS_SECTION_RE = re.compile(
    r"(吸入|接触眼睛|接触皮肤|误服|食入|灭火|泄漏|泄露|收集|废弃处置|安全存储|储存条件|急救)")

# S2 危险性说明 / 防范说明 代码规则
_H_RE = re.compile(r"H\d{3}")
_P_RE = re.compile(r"P\d{3}")


@dataclass
class SectionCheck:
    """单节判定结果."""
    num: int
    hard_missing: list[str] = dc_field(default_factory=list)   # 硬性缺失字段
    soft_missing: list[str] = dc_field(default_factory=list)   # 软性缺失字段
    vetoes: list[str] = dc_field(default_factory=list)         # 否决项
    ok: bool = True

    @property
    def level(self) -> str:
        if self.vetoes or self.hard_missing:
            return "fail"
        if self.soft_missing:
            return "warn"
        return "pass"


# PEA-4139 模板自身的软性缺失 (模板都没写的字段, 不算合格障碍).
# 实测: 模板判定为 warn, 软缺正好这 9 个字段 → "完美符合 PEA-4139" 的基准
# 即: 软缺 ⊆ 本集合 且 无硬缺/无否决 = 完美合格, 可原样入标准库.
TEMPLATE_SOFT_GAP: frozenset = frozenset({
    (1, "产品名称"), (1, "应急电话"), (1, "电子邮件"),
    (2, "标签有害成分"),
    (3, "组分"), (3, "CAS号"),
    (9, "相对密度"), (9, "熔点/凝固点"),
    (15, "其它的规定"),
})


@dataclass
class ComplianceReport:
    """整个 MSDS 的合格判定."""
    file_name: str
    sections: dict[int, SectionCheck] = dc_field(default_factory=dict)
    missing_sections: list[int] = dc_field(default_factory=list)

    @property
    def level(self) -> str:
        if self.missing_sections:
            return "fail"
        secs = self.sections.values()
        if any(s.level == "fail" for s in secs):
            return "fail"
        if any(s.level == "warn" for s in secs):
            return "warn"
        return "pass"

    @property
    def soft_gaps(self) -> set:
        """全部软缺 (节, 字段) 集合."""
        return {(num, s) for num, c in self.sections.items()
                for s in c.soft_missing}

    @property
    def perfect(self) -> bool:
        """完美合格: 无硬缺/无否决, 且软缺不超过 PEA-4139 模板自身缺失.

        判定标准以模板为准 (模板自身软缺即 TEMPLATE_SOFT_GAP):
        软缺 ⊆ 模板软缺基线 → 与模板同一水平, 可原样入标准库.
        """
        return (self.level != "fail"
                and self.soft_gaps <= TEMPLATE_SOFT_GAP)

    def summary(self) -> dict:
        return {
            "file": self.file_name,
            "level": self.level,
            "sections": len(self.sections),
            "missing_sections": self.missing_sections,
            "hard_missing": [f"S{s.num}:{','.join(s.hard_missing)}"
                             for s in self.sections.values() if s.hard_missing],
            "soft_missing": sum(len(s.soft_missing)
                                for s in self.sections.values()),
            "vetoes": [f"S{s.num}:{v}" for s in self.sections.values()
                       for v in s.vetoes],
        }


# 每节: (hard 字段列表, soft 字段列表). 字段名 = Schema 标准名 (见 schema.py).
_SECTION_STANDARD: dict[int, tuple[list[str], list[str]]] = {
    0: (["Version", "产品名称", "公司名称", "产品型号", "修订日期"], []),
    1: (["产品使用建议和使用限制", "供应商名称", "供应商地址",
         "电话", "传真"],
        ["产品名称", "中文名称", "产品类型", "应急电话", "电子邮件"]),
    2: (["GHS危险性类别", "信号词", "危害性说明", "防范说明"],
        ["象形图", "标签有害成分", "其他危险"]),
    3: (["产品类型"], ["组分", "CAS号"]),
    4: (["一般措施", "误服", "接触眼睛", "接触皮肤", "吸入"], []),
    5: (["合适的灭火剂", "不合适的灭火剂", "物质或混合物的特殊危害",
         "消防预防措施和保护设备"], []),
    6: (["个人预防措施、应急程序", "环境保护措施",
         "污染物收集和清除的方法"], []),
    7: (["安全操作防范", "安全储存条件"], []),
    8: (["呼吸系统防护", "手部防护", "眼睛防护", "身体防护"],
        ["防护手套的合适材料", "建议"]),
    9: (["外观"], ["pH值", "闪点", "密度", "相对密度", "水溶性", "离子性",
                   "初沸点/沸点", "熔点/凝固点", "粘度", "其他信息",
                   "饱和蒸气压", "蒸气密度", "分配系数", "自燃温度",
                   "分解温度", "易燃性", "固体含量", "蒸发速率"]),
    10: (["稳定性", "危险的分解产物", "危险反应"], []),
    11: (["急性毒性"], ["主要皮肤刺激性", "主要眼睛刺激性", "致敏性",
                         "致突变性", "致癌性", "生殖毒性",
                         "特异性靶器官系统毒性", "吸入危险", "附加信息"]),
    12: (["生态毒性"], ["持久性和降解性", "其他不利的影响"]),
    13: (["处理方法"], []),
    14: (["公路和铁路运输", "海上运输", "空运"], ["用户特殊注意事项"]),
    15: (["符合下列法规要求"], ["其它的规定"]),
    16: ([], []),   # 仅 note, 无必需字段
}


def _field_names(num: int) -> set[str]:
    """该节 Schema 定义的标准字段名集合."""
    return {f.name for f in SECTION_SCHEMAS.get(num, _BASE_SCHEMAS.get(num, []))}


def _section_values(result, num: int) -> dict[str, list[str]]:
    """该节归一化后的 {标准字段名: [原始值, ...]}, 仅 field 行."""
    out: dict[str, list[str]] = {}
    sec = result.sections.get(num)
    if not sec:
        return out
    std_names = _field_names(num)
    for row in sec.iter_rows():
        if row.kind != "field":
            continue
        name = standard_name(num, row.label)
        if not name:
            continue
        out.setdefault(name, []).append(row.value)
    return out


def _veto_s2(sec) -> list[str]:
    """S2 否决项: 跨节内容 / H·P 代码缺失 / 标签粘连."""
    vetoes: list[str] = []
    labels = [r.label for r in sec.iter_rows()
              if r.kind == "field" and r.label.strip()]
    for lbl in labels:
        if _S2_CROSS_SECTION_RE.search(lbl):
            vetoes.append(f"跨节内容「{lbl[:15]}」混入S2")
    # H / P 代码
    for r in sec.iter_rows():
        if r.kind != "field":
            continue
        name = standard_name(2, r.label)
        val = r.value or ""
        if name == "危害性说明" and val and not _H_RE.search(val):
            vetoes.append("危害性说明有值但无H代码")
        if name == "防范说明" and val and not _P_RE.search(val):
            vetoes.append("防范说明有值但无P代码")
    return vetoes


def _veto_s9(sec) -> list[str]:
    """S9 否决项: 值混入标签.

    仅判定"值形态"标签 (含数字/数值单位/百分比/区间 等), 如
      '乳白色液体' / '7-9 （按 1' / '约1.06g/cm3可混溶' / '不适用无数据'
    不含数字的标签变体 (如 'pH 值'/'辛醇/ 水分配系数的对数值') 是字段
    写法, 由 standard_name 归一化, 不算值混入.
    """
    vetoes: list[str] = []
    _value_like = re.compile(r"[\d.]|g/cm|%|液体|固体|无数据|不适用")
    for r in sec.iter_rows():
        if r.kind != "field":
            continue
        name = standard_name(9, r.label)
        if name in _field_names(9):
            continue
        lbl = r.label.strip()
        if lbl and _value_like.search(lbl):
            vetoes.append(f"S9标签疑似值混入「{lbl[:15]}」")
    return vetoes


def _is_no_hazard(sec) -> bool:
    """S2 是否"无危险豁免": GHS危险性类别 值为 根据GHS不属于危险物/无危险 等."""
    for r in sec.iter_rows():
        if r.kind != "field":
            continue
        if standard_name(2, r.label) == "GHS危险性类别":
            v = (r.value or "").strip()
            if any(k in v for k in ("不属于危险", "无危险", "非危险品", "不属于危害")):
                return True
    return False


def _check_section(result, num: int, hard: list[str],
                   soft: list[str]) -> SectionCheck:
    sec = result.sections.get(num)
    check = SectionCheck(num=num)
    if sec is None:
        check.hard_missing = list(hard) or ["<节缺失>"]
        check.ok = False
        return check

    # S2 无危险豁免: 该豁免字段免检 (无危险品无 信号词/危险性说明/防范说明)
    waived: set[str] = set()
    if num == 2 and _is_no_hazard(sec):
        waived = {"信号词", "危害性说明", "防范说明", "象形图", "标签有害成分"}

    vals = _section_values(result, num)
    for f in hard:
        if f in waived:
            continue
        v = vals.get(f, [])
        if not any(x.strip() for x in v):
            check.hard_missing.append(f)
    for f in soft:
        v = vals.get(f, [])
        if not any(x.strip() for x in v):
            check.soft_missing.append(f)

    # 节专用否决项
    if num == 2:
        check.vetoes = _veto_s2(sec)
    if num == 9:
        check.vetoes = _veto_s9(sec)
    check.ok = not check.vetoes and not check.hard_missing
    return check


def evaluate(result) -> ComplianceReport:
    """对单个 ParseResult 做合格判定."""
    rep = ComplianceReport(file_name=result.file_name)
    for num in range(0, 17):
        hard, soft = _SECTION_STANDARD.get(num, ([], []))
        check = _check_section(result, num, hard, soft)
        rep.sections[num] = check
    # 缺失节 (S0-S15 必须齐全)
    present = set(result.sections.keys())
    for num in range(0, 16):
        if num not in present:
            rep.missing_sections.append(num)
    return rep


def evaluate_dir(in_dir: str | Path, min_sections: int = 8) -> dict:
    """全库合格筛选: {perfect, warn, fail, not_msds, failed, files}.

    min_sections: 低于该检出节数的文件视为非16节MSDS (如 TDS),
    归入 not_msds 而非 fail, 不参与标准库筛选.

    perfect: 完美合格 (硬缺0 + 否决0 + 软缺 ⊆ 模板软缺基线),
    即与 PEA-4139 模板同一水平的文件, 可直接入标准库.
    """
    in_dir = Path(in_dir)
    files = sorted(p for p in in_dir.rglob("*.docx")
                   if not p.name.startswith("~$"))
    result = {"perfect": [], "warn": [], "fail": [], "not_msds": [],
              "failed": [], "files": len(files)}
    for p in files:
        try:
            r = read_msds(p)
            rep = evaluate(r)
            if len(r.sections) < min_sections:
                result["not_msds"].append({
                    "file": p.name, "path": str(p),
                    "sections": len(r.sections),
                })
                continue
            row = {
                "file": p.name,
                "path": str(p),
                "level": rep.level,
                "missing": rep.missing_sections,
                "hard": len(rep.summary()["hard_missing"]),
                "soft": rep.summary()["soft_missing"],
                "vetoes": len(rep.summary()["vetoes"]),
                "perfect": rep.perfect,
            }
            bucket = "perfect" if rep.perfect else rep.level
            result[bucket].append(row)
        except Exception as exc:
            result["failed"].append(f"{p.name} ({exc.__class__.__name__})")
    return result

