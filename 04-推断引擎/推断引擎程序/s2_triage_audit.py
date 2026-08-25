"""Read-only large-scale Section 2 three-way audit.

The audit keeps three sources separate:

1. ``original``: Section 2 values already stored for the unique model in the
   formal Chinese database (and the source Word path when it is available).
2. ``legal_manual``: an independent, conservative legal review.  It only
   asserts a result where S1/S3/S9 contain a direct, uniquely classifiable
   fact.  For the current first pass that means the GB 30000.7-2013 flash
   point rule for flammable liquids.  Component health/environment hazards
   are not asserted from a CAS number alone because the current CAS database
   has no traceable H-code + clause fact table.
3. ``program``: the existing ``collect_fact_package`` + ``infer_s2`` path.

No formal database or source document is written.  The output directory is a
new verification artifact and can be removed without changing production
data.
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import sqlite3
import sys
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

ENGINE_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = ENGINE_DIR.parents[1]
GUI_ROOT = PROJECT_ROOT / "02-检索系统"
MODEL_DB = PROJECT_ROOT / "03-数据库" / "正式库" / "Data Base" / "msds_standard.db"
CAS_DB = PROJECT_ROOT / "03-数据库" / "正式库" / "Data Base" / "cas_library.db"
RESULT_LIBRARY = ENGINE_DIR / "s2_result_library.json"
if str(ENGINE_DIR) not in sys.path:
    sys.path.insert(0, str(ENGINE_DIR))
if str(GUI_ROOT) not in sys.path:
    sys.path.insert(0, str(GUI_ROOT))

from s2_fact_package import collect_fact_package  # noqa: E402
from s2_inference import infer_s2  # noqa: E402

try:  # Direct Word comparison is useful but not allowed to block the audit.
    from core.docx_reader import read_msds  # type: ignore  # noqa: E402
except Exception:  # pragma: no cover - environment diagnostic is reported
    read_msds = None


FIELD_NAMES = (
    "GHS危险性类别",
    "象形图",
    "信号词",
    "危险性说明",
    "防范说明",
    "物理和化学危险",
    "健康危害",
    "环境危害",
    "其他危害",
)
H_RE = re.compile(r"H\s*(\d{3})", re.IGNORECASE)
NUMBER_RE = re.compile(r"[-+]?\d+(?:[.,]\d+)?")

LEGAL_SOURCES = [
    {
        "id": "GB30000.1-2024",
        "title": "化学品分类和标签规范 第1部分：通则",
        "status": "现行；2025-08-01实施",
        "official_url": "https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=C008AEDBFD9A16F3C5BEB671C20618DD",
        "local_path": str(ENGINE_DIR / "法规匹配库" / "08_未分类_待核验" / "GB30000.1-2024.pdf"),
        "role": "混合物临界值、GHS分类与SDS信息边界；当前人工判定的总则",
    },
    {
        "id": "GB30000.7-2013",
        "title": "化学品分类和标签规范 第7部分：易燃液体",
        "status": "现行",
        "official_url": "https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=3FCEB3CB81076CDB311D0D3F2783B2F7",
        "local_path": str(ENGINE_DIR / "法规匹配库" / "03_技术标准_分类与框架" / "01_GHS分类_GB30000系列" / "GB 30000.7-2013.pdf"),
        "role": "闪点/初沸点到易燃液体类别、H224/H225/H226/H227、信号词和象形图",
    },
    {
        "id": "GB/T17519-2013",
        "title": "化学品安全技术说明书编写指南",
        "status": "现行",
        "official_url": "https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=0DF24B277B29E2F0B39D13D0F2B7BA07",
        "local_path": str(ENGINE_DIR / "法规匹配库" / "08_未分类_待核验" / "GB-T17519-2013.pdf"),
        "role": "Section 2分类、标签要素、混合物健康/环境组分浓度限值及未知信息表达",
    },
    {
        "id": "GB/T16483-2008",
        "title": "化学品安全技术说明书 内容和项目顺序",
        "status": "现行",
        "official_url": "https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=67992CA972A4CF9222095CA06064724A&refer=outter",
        "local_path": str(ENGINE_DIR / "法规匹配库" / "08_未分类_待核验" / "GB-T16483-2008.pdf"),
        "role": "SDS 16节结构及Section 2边界",
    },
    {
        "id": "GB15258-2009",
        "title": "化学品安全标签编写规定",
        "status": "现行",
        "official_url": "https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=4D487D68BF0BD87E68CE0EA68183DAD6",
        "local_path": str(ENGINE_DIR / "法规匹配库" / "01_法律法规_国内" / "03_部门规章" / "GB 15258-2009 化学品安全标签编写规定.txt"),
        "role": "标签要素的编排和一致性要求；不替代分类标准",
    },
    {
        "id": "危险化学品安全管理条例",
        "title": "危险化学品安全管理条例",
        "status": "国务院行政法规；用于SDS/标签相符性义务",
        "official_url": "https://www.samr.gov.cn/zw/zfxxgk/fdzdgknr/bgt/art/2023/art_7f7163c7940b4c4f8639c6f6c3312ede.html",
        "local_path": str(ENGINE_DIR / "法规匹配库" / "01_法律法规_国内" / "01_国家法律" / "危险化学品安全管理条例.txt"),
        "role": "要求SDS和安全标签与危险化学品相符并符合国家标准；不是GHS分类计算表",
    },
]


def _text(value: Any) -> str:
    if isinstance(value, list):
        return "；".join(_text(item) for item in value if _text(item))
    return str(value or "").strip()


def _compact(value: Any) -> str:
    text = _text(value).replace("％", "%").replace("℃", "°C")
    text = re.sub(r"[\s\u3000]+", "", text)
    return text


def _json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True)


def _db_rows(connection: sqlite3.Connection, sql: str, params: Iterable[Any] = ()) -> list[dict[str, Any]]:
    connection.row_factory = sqlite3.Row
    return [dict(row) for row in connection.execute(sql, tuple(params)).fetchall()]


def _db_original_s2(connection: sqlite3.Connection, model: str) -> dict[str, Any]:
    rows = _db_rows(
        connection,
        """
        SELECT f.seq, f.label, f.std_name, f.value, f.kind, f.editable
        FROM msds_field f JOIN msds_model m ON m.model_id = f.model_id
        WHERE m.model = ? AND f.section = 2
        ORDER BY f.id
        """,
        [model],
    )
    values: dict[str, Any] = {}
    for row in rows:
        name = str(row.get("std_name") or row.get("label") or "未命名字段")
        value = row.get("value") or "无数据"
        if name not in values:
            values[name] = value
        elif not isinstance(values[name], list):
            values[name] = [values[name], value]
        else:
            values[name].append(value)
    return {"values": values, "rows": rows}


def _word_s2_text(path: str) -> tuple[str, str]:
    if read_msds is None:
        return "", "word_reader_unavailable"
    source = Path(path)
    if not source.exists():
        return "", "source_missing"
    try:
        parsed = read_msds(source)
        section = parsed.sections.get(2)
        if section is None:
            return "", "section_2_missing"
        chunks: list[str] = []
        for field in section.fields:
            chunks.append(f"{field.label}：{field.value}")
        chunks.extend(section.lines)
        return "\n".join(chunk for chunk in chunks if chunk).strip(), "ok"
    except Exception as exc:  # keep one corrupt source from stopping all models
        return "", f"word_parse_error:{type(exc).__name__}:{exc}"


def _first_fact(facts: dict[str, Any], *names: str) -> str:
    for name in names:
        if name in facts and _text(facts[name]):
            return _text(facts[name])
    return ""


def _temperature(raw: Any) -> dict[str, Any] | None:
    text = _text(raw).replace("℃", "").replace("°C", "").replace("ºC", "")
    if not text or text in {"无数据", "不适用", "无闪点", "初沸点下无闪点"}:
        return None
    match = NUMBER_RE.search(text.replace("，", ","))
    if not match:
        return None
    try:
        value = float(match.group(0).replace(",", "."))
    except ValueError:
        return None
    before = text[: match.start()].strip()
    operator = "exact"
    if re.search(r"(?:大于|高于|超过|>|≥|不低于)", before):
        operator = "gt_or_equal" if "≥" in before or "不低于" in before else "gt"
    elif re.search(r"(?:小于|低于|不超过|≤|<)", before):
        operator = "lt_or_equal" if "≤" in before or "不超过" in before else "lt"
    elif re.search(r"(?:约|大约|左右|接近)", before):
        operator = "approx"
    return {"raw": text, "value": value, "operator": operator}


def _contains_liquid(facts: dict[str, Any]) -> bool:
    appearance = _first_fact(facts, "外观")
    product_type = _first_fact(facts, "产品类型")
    return any(token in appearance + product_type for token in ("液", "乳液", "溶液", "分散体")) or bool(
        _first_fact(facts, "闪点")
    )


def _manual_flammable(s9_facts: dict[str, Any]) -> dict[str, Any]:
    flash_raw = _first_fact(s9_facts, "闪点")
    boil_raw = _first_fact(s9_facts, "初沸点")
    flash = _temperature(flash_raw)
    boil = _temperature(boil_raw)
    base = ["GB30000.1-2024 4.1.1/6.1.1", "GB30000.7-2013 4.2表1、附录B/C/D"]
    unresolved = {
        "status": "not_uniquely_determined",
        "categories": [],
        "hazard_codes": [],
        "pictogram_codes": [],
        "signal_word": "无法唯一确定",
        "hazard_statements": [],
        "basis": base,
        "reason": "S9未提供可直接落入唯一易燃液体类别的闪点/初沸点证据。",
        "flash_raw": flash_raw or "无数据",
        "boil_raw": boil_raw or "无数据",
    }
    if not flash or not _contains_liquid(s9_facts):
        return unresolved
    f, op = flash["value"], flash["operator"]
    # “初沸点下无闪点” is a negative test statement, not a temperature.
    if "无闪点" in flash_raw:
        return unresolved
    category = code = pict = signal = statement = ""
    if op in {"gt", "gt_or_equal"} and f > 93:
        return {
            **unresolved,
            "status": "direct_no_flammable_classification",
            "signal_word": "无",
            "reason": "闪点明确高于93℃，不落入GB 30000.7-2013易燃液体定义。",
        }
    if op in {"lt", "lt_or_equal"} and f < 23:
        # For a flash point below 23℃, initial boiling point alone selects 1/2.
        if not boil:
            return {**unresolved, "reason": "闪点低于23℃，但缺少初沸点，不能在类别1/2之间选择。"}
        if boil["value"] <= 35:
            category, code, pict, signal, statement = "易燃液体，类别 1", "H224", "GHS02", "危险", "极易燃液体和蒸气"
        elif boil["value"] > 35:
            category, code, pict, signal, statement = "易燃液体，类别 2", "H225", "GHS02", "危险", "高度易燃液体和蒸气"
    elif op in {"gt", "gt_or_equal"} and f >= 23:
        if f > 60 and f <= 93:
            category, code, pict, signal, statement = "易燃液体，类别 4", "H227", "无图形符号", "警告", "可燃液体"
        elif f <= 60:
            category, code, pict, signal, statement = "易燃液体，类别 3", "H226", "GHS02", "警告", "易燃液体和蒸气"
    elif op == "exact":
        if f < 23:
            if not boil:
                return {**unresolved, "reason": "闪点低于23℃，但缺少初沸点，不能在类别1/2之间选择。"}
            if boil["value"] <= 35:
                category, code, pict, signal, statement = "易燃液体，类别 1", "H224", "GHS02", "危险", "极易燃液体和蒸气"
            elif boil["value"] > 35:
                category, code, pict, signal, statement = "易燃液体，类别 2", "H225", "GHS02", "危险", "高度易燃液体和蒸气"
        elif f <= 60:
            category, code, pict, signal, statement = "易燃液体，类别 3", "H226", "GHS02", "警告", "易燃液体和蒸气"
        elif f <= 93:
            category, code, pict, signal, statement = "易燃液体，类别 4", "H227", "无图形符号", "警告", "可燃液体"
        else:
            return {
                **unresolved,
                "status": "direct_no_flammable_classification",
                "signal_word": "无",
                "reason": "闪点高于93℃，不落入GB 30000.7-2013易燃液体定义。",
            }
    else:
        return {**unresolved, "reason": "闪点含约数/不等式，当前表达不足以唯一落入边界类别。"}
    if not category:
        return unresolved
    return {
        **unresolved,
        "status": "directly_determined",
        "categories": [category],
        "hazard_codes": [code],
        "pictogram_codes": [] if pict == "无图形符号" else [pict],
        "pictogram_display": pict,
        "signal_word": signal,
        "hazard_statements": [f"{code} {statement}"],
        "reason": "S9闪点/初沸点直接命中GB 30000.7-2013分类阈值。",
    }


def _manual_legal(package: dict[str, Any]) -> dict[str, Any]:
    s9 = ((package.get("input_snapshot") or {}).get("s9") or {}).get("facts") or {}
    usages = list((package.get("substance_snapshot") or {}).get("usage") or [])
    physical = _manual_flammable(s9)
    cas_values = [str(row.get("cas_no") or "").strip() for row in usages if str(row.get("cas_no") or "").strip()]
    confidential = [row for row in usages if not str(row.get("cas_no") or "").strip()]
    reasons = [
        "当前CAS库仅提供物质/别名/含量/监管状态，没有可追溯的‘CAS-H码-适用法规条款’事实表。",
        "GB/T17519-2013 3.2.2要求危险性分类采用相关国家标准或已统一分类目录；不能从CAS号或型号名本身推定健康/环境分类。",
    ]
    if confidential:
        reasons.append(f"存在{len(confidential)}条无CAS/保密组分，无法对其健康/环境危险性作唯一法规判定。")
    elif not cas_values:
        reasons.append("S3/CAS库未形成可核验的明确CAS组分。")
    else:
        reasons.append(f"已识别CAS {len(cas_values)}条，但其产品级混合物分类仍缺少可追溯法规分类事实。")
    if physical["status"] == "directly_determined":
        status = "partial_direct_only"
    elif physical["status"] == "direct_no_flammable_classification":
        status = "partial_direct_no_flammable_only"
    else:
        status = "insufficient_legal_evidence"
    return {
        "status": status,
        "categories": physical["categories"],
        "hazard_codes": physical["hazard_codes"],
        "pictogram_codes": physical["pictogram_codes"],
        "pictogram_display": physical.get("pictogram_display", "无数据"),
        "signal_word": physical["signal_word"],
        "hazard_statements": physical["hazard_statements"],
        "health_hazard": "无法由现有唯一法规事实源唯一确定",
        "environment_hazard": "无法由现有唯一法规事实源唯一确定",
        "basis": physical["basis"],
        "reason": "；".join([physical["reason"], *reasons]),
        "cas_count": len(cas_values),
        "confidential_count": len(confidential),
        "flash_raw": physical["flash_raw"],
        "boil_raw": physical["boil_raw"],
    }


def _h_codes(value: Any) -> list[str]:
    seen: list[str] = []
    for match in H_RE.finditer(_text(value).upper()):
        code = f"H{match.group(1)}"
        if code not in seen:
            seen.append(code)
    return seen


def _original_semantics(original: dict[str, Any]) -> dict[str, Any]:
    values = original["values"]
    combined = "；".join(_text(values.get(name, "")) for name in FIELD_NAMES)
    compact = _compact(combined)
    no_hazard = "根据GHS不属于危险物" in combined or "无危险" in combined
    return {
        "hazard_codes": _h_codes(combined),
        "pictogram_text": _text(values.get("象形图", "无数据")),
        "no_hazard_claim": no_hazard,
        "status": "explicit_no_hazard_claim" if no_hazard else "raw_result_present" if compact else "empty_result",
    }


def _program_hazard_codes(result: dict[str, Any]) -> list[str]:
    """Recover the program's internal H-code evidence without trusting text.

    The current engine exposes known-substance H codes in condition results,
    while its human-facing hazard statement values omit the code prefix.  The
    audit therefore reads the condition trace and adds the independent
    flammable-liquid mapping from the condition result.
    """
    category_to_code = {
        "易燃液体，类别 1": "H224",
        "易燃液体，类别 2": "H225",
        "易燃液体，类别 3": "H226",
        "易燃液体，类别 4": "H227",
    }
    codes: list[str] = []
    for condition in result.get("conditions") or []:
        condition_result = condition.get("result")
        if isinstance(condition_result, dict):
            candidates = condition_result.get("hazard_codes") or []
            for code in candidates:
                code = str(code).upper().replace(" ", "")
                if code.startswith("H") and code not in codes:
                    codes.append(code)
        elif condition.get("id") == "flammable_liquid_by_s9":
            code = category_to_code.get(str(condition_result or "").strip())
            if code and code not in codes:
                codes.append(code)
        code = str((condition.get("inputs") or {}).get("code") or "").upper().replace(" ", "")
        if re.fullmatch(r"H\d{3}", code) and condition.get("matched") and code not in codes:
            codes.append(code)
    for value in (result.get("result") or {}).values():
        for code in _h_codes(value):
            if code not in codes:
                codes.append(code)
    return codes


def _program_semantics(result: dict[str, Any]) -> dict[str, Any]:
    return {
        "hazard_codes": _program_hazard_codes(result),
        "pictogram_codes": list(result.get("candidate_label_elements", {}).get("pictogram_codes") or []),
        "categories": _text(result.get("result", {}).get("GHS危险性类别", "无数据")),
        "pictogram_display": _text(result.get("result", {}).get("象形图", "无数据")),
        "signal_word": _text(result.get("result", {}).get("信号词", "无数据")),
    }


def _compare(original: dict[str, Any], legal: dict[str, Any], program: dict[str, Any]) -> tuple[str, str]:
    manual_codes = set(legal["hazard_codes"])
    program_codes = set(program["hazard_codes"])
    unsupported = sorted(program_codes - manual_codes)
    manual_physical = set(legal["categories"])
    program_physical = {part.strip() for part in program["categories"].split("；") if "易燃液体" in part}
    notes: list[str] = []
    if legal["status"] == "partial_direct_only":
        if program_physical != manual_physical:
            notes.append(f"直接物理分类不一致：法规人工={sorted(manual_physical)}；程序={sorted(program_physical)}")
        # Only compare label elements when the program actually produced a
        # label-bearing hazard result.  ``无数据`` is an unresolved program
        # state, not evidence that the legal result is wrong.
        if legal["pictogram_codes"] != program["pictogram_codes"]:
            notes.append(f"标签象形图不一致：法规人工={legal['pictogram_codes'] or ['无图形符号']}；程序={program['pictogram_codes'] or ['无图形符号']}")
        if program_codes and legal["signal_word"] != program["signal_word"]:
            notes.append(f"信号词不一致：法规人工={legal['signal_word']}；程序={program['signal_word']}")
    elif legal["status"] == "partial_direct_no_flammable_only":
        if program_codes or program_physical or program["pictogram_codes"]:
            notes.append("法规人工已由S9直接排除易燃液体分类，但程序仍输出了物理/标签危险结果")
    if unsupported:
        notes.append("程序输出了当前唯一法规事实源未支持的健康/环境或组分危险码：" + ",".join(unsupported))
    if original["no_hazard_claim"] and (legal["categories"] or unsupported):
        notes.append("原始S2声称‘根据GHS不属于危险物’，但三方证据至少存在直接物理分类或未闭合健康/环境证据")
    if not notes:
        if legal["status"] == "insufficient_legal_evidence":
            return "cannot_conclude", "法规事实源不足，程序未形成可核验结论。"
        if legal["status"] == "partial_direct_no_flammable_only":
            return "aligned_on_supported_facts", "S9直接排除易燃液体分类；健康/环境分类仍需补充物质级法规事实。"
        return "aligned_on_supported_facts", "在当前可唯一判定的事实范围内一致。"
    if unsupported:
        return "program_overreach", "；".join(notes)
    if any("象形图" in n or "信号词" in n or "直接物理分类" in n for n in notes):
        return "program_label_or_physical_mismatch", "；".join(notes)
    return "original_vs_legal_conflict", "；".join(notes)


def _row_for_model(model: str, model_row: dict[str, Any], model_con: sqlite3.Connection) -> dict[str, Any]:
    package = collect_fact_package(model)
    program_result = infer_s2(package)
    original = _db_original_s2(model_con, model)
    legal = _manual_legal(package)
    original_semantics = _original_semantics(original)
    program_semantics = _program_semantics(program_result)
    mismatch_class, mismatch_detail = _compare(original_semantics, legal, program_semantics)
    word_text, word_status = _word_s2_text(str(model_row.get("source_file") or ""))
    s1 = ((package.get("input_snapshot") or {}).get("s1") or {}).get("facts") or {}
    s3 = ((package.get("input_snapshot") or {}).get("s3") or {}).get("facts") or {}
    s9 = ((package.get("input_snapshot") or {}).get("s9") or {}).get("facts") or {}
    usages = list((package.get("substance_snapshot") or {}).get("usage") or [])
    return {
        "model": model,
        "source_file": str(model_row.get("source_file") or ""),
        "source_sha256": str(model_row.get("sha256") or ""),
        "word_source_status": word_status,
        "word_section2_raw": word_text,
        "s1_facts": _json(s1),
        "s3_facts": _json(s3),
        "s9_facts": _json(s9),
        "s9_flash": _first_fact(s9, "闪点") or "无数据",
        "s9_initial_boiling_point": _first_fact(s9, "初沸点") or "无数据",
        "cas_usage_count": len(usages),
        "cas_usage": _json(usages),
        "original_s2": _json(original["values"]),
        "original_s2_status": original_semantics["status"],
        "original_s2_h_codes": _json(original_semantics["hazard_codes"]),
        "original_s2_no_hazard_claim": original_semantics["no_hazard_claim"],
        "legal_manual_status": legal["status"],
        "legal_manual_categories": _json(legal["categories"]),
        "legal_manual_h_codes": _json(legal["hazard_codes"]),
        "legal_manual_pictograms": _json(legal["pictogram_codes"]),
        "legal_manual_pictogram_display": legal["pictogram_display"],
        "legal_manual_signal_word": legal["signal_word"],
        "legal_manual_hazard_statements": _json(legal["hazard_statements"]),
        "legal_manual_health_hazard": legal["health_hazard"],
        "legal_manual_environment_hazard": legal["environment_hazard"],
        "legal_manual_basis": _json(legal["basis"]),
        "legal_manual_reason": legal["reason"],
        "legal_cas_count": legal["cas_count"],
        "confidential_component_count": legal["confidential_count"],
        "program_status": str(program_result.get("status") or ""),
        "program_categories": program_semantics["categories"],
        "program_h_codes": _json(program_semantics["hazard_codes"]),
        "program_pictogram_codes": _json(program_semantics["pictogram_codes"]),
        "program_pictogram_display": program_semantics["pictogram_display"],
        "program_signal_word": program_semantics["signal_word"],
        "program_review_reasons": _json(program_result.get("review_reasons") or []),
        "program_conditions": _json(program_result.get("conditions") or []),
        "mismatch_class": mismatch_class,
        "mismatch_detail": mismatch_detail,
        "program_rule_version": str(program_result.get("rule_version") or ""),
    }


def _write_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    if not rows:
        return
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)


def _write_xlsx(path: Path, rows: list[dict[str, Any]], summary: dict[str, Any]) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    wb = Workbook()
    ws = wb.active
    ws.title = "汇总"
    for key, value in summary.items():
        ws.append([key, value if not isinstance(value, (dict, list)) else _json(value)])
    ws.append([])
    ws.append(["三方对照明细"])
    headers = list(rows[0]) if rows else []
    ws.append(headers)
    for row in rows:
        ws.append([row.get(header, "") for header in headers])
    detail = wb.create_sheet("人工复核队列")
    queue_headers = ["model", "mismatch_class", "mismatch_detail", "legal_manual_status", "legal_manual_reason", "program_categories", "program_h_codes", "original_s2"]
    detail.append(queue_headers)
    for row in rows:
        if row.get("mismatch_class") != "aligned_on_supported_facts":
            detail.append([row.get(k, "") for k in queue_headers])
    source = wb.create_sheet("法规源登记")
    source_headers = list(LEGAL_SOURCES[0])
    source.append(source_headers)
    for item in LEGAL_SOURCES:
        source.append([item.get(k, "") for k in source_headers])
    for sheet in wb.worksheets:
        sheet.freeze_panes = "A2"
        sheet.sheet_view.showGridLines = False
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
        for row in sheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
        for column in sheet.columns:
            letter = column[0].column_letter
            sheet.column_dimensions[letter].width = min(60, max(14, max(len(str(c.value or "")) for c in column[:30]) + 2))
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def _write_report(path: Path, rows: list[dict[str, Any]], summary: dict[str, Any]) -> None:
    classes = Counter(row["mismatch_class"] for row in rows)
    lines = [
        "# Section 2 三方对照验证报告",
        "",
        f"生成时间：{summary['generated_at']}",
        f"样本型号数：{summary['model_count']}",
        "",
        "## 对照定义",
        "",
        "- 原有数据结果：正式中文总型号库 Section 2 原值，同时保留型号对应的 Word 来源状态。",
        "- 法规人工判定：独立于 `s2_inference.py`，只对 S1/S3/S9 中可由现行标准唯一判定的事实给出结论；当前首轮可直接判定的是 GB 30000.7-2013 易燃液体规则。健康/环境分类若缺少可追溯的 CAS-H码-法规条款事实，明确标为不能唯一确定。",
        "- 程序推导结果：现有 `s2_fact_package.py` + `s2_inference.py` 的实际输出。",
        "",
        "## 汇总",
        "",
        f"- 原始 Section 2 有内容：{summary['original_nonempty_count']}；空或无结构结果：{summary['original_empty_count']}",
        f"- 原始 S2 明确写‘根据GHS不属于危险物’：{summary['original_no_hazard_claim_count']}；原始记录含H码：{summary['original_h_code_count']}",
        f"- 法规人工直接判定易燃液体类别：{summary['manual_direct_count']}；直接排除易燃液体分类：{summary['manual_direct_no_flammable_count']}；健康/环境等法规证据不足：{summary['manual_insufficient_count']}",
        f"- CAS库物质数：{summary['cas_substance_count']}；程序结果库已编码物质数：{summary['program_known_substance_count']}",
        f"- 程序状态：{_json(summary['program_status'])}",
        f"- 差异分类：{_json(classes)}",
        f"- 直接物理分类不一致：{summary['direct_physical_mismatch_count']}；直接标签象形图不一致：{summary['direct_pictogram_mismatch_count']}；程序自动通过但仍有差异：{summary['auto_pass_mismatch_count']}",
        f"- 程序越过当前法规事实源的H码频次：{_json(summary['unsupported_h_code_frequency'])}",
        "",
        "## 当前确定的问题",
        "",
        "1. GB 30000.7-2013 附录B/C/D明确：易燃液体类别4为‘无图形符号、警告、H227可燃液体’。程序当前把 H227 映射为 GHS02，并由此形成标签差异。",
        "2. 现有 CAS 库没有可追溯的物质级 H 码、分类类别、适用标准条款字段；程序结果库只有 12 个已编码物质，不能代表完整 CAS 库的法律真值。",
        "3. 大量原始 S2 直接写‘根据GHS不属于危险物’，但 GB/T 17519-2013 和 GB 30000.1-2024要求未知资料与已证实不存在危险区分；当输入存在保密组分、CAS未闭合或直接物理危险时，该表述不能作为法规真值。",
        "4. 程序会将 CAS 结果库中的健康/环境 H 码直接合并到产品级结论；在没有产品级混合物阈值所需完整证据时，这些结果必须进入人工复核，不能自动通过。",
        "",
        "## 复核边界",
        "",
        "本报告不是对所有健康/环境危险类别的最终法律意见。它已经把不能从现有唯一事实源唯一推导的项目列入人工复核队列；下一轮应补充 CAS-H码-法规条款事实表，再重新跑三方对照。",
    ]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def run(out_dir: Path) -> dict[str, Any]:
    out_dir.mkdir(parents=True, exist_ok=True)
    if not MODEL_DB.exists() or not CAS_DB.exists():
        raise FileNotFoundError(f"正式数据库缺失: {MODEL_DB} / {CAS_DB}")
    with sqlite3.connect(MODEL_DB) as model_con:
        model_con.row_factory = sqlite3.Row
        model_rows = [dict(row) for row in model_con.execute("SELECT * FROM msds_model ORDER BY model").fetchall()]
        rows: list[dict[str, Any]] = []
        errors: list[dict[str, str]] = []
        for item in model_rows:
            model = str(item["model"])
            try:
                rows.append(_row_for_model(model, item, model_con))
            except Exception as exc:
                errors.append({"model": model, "error": f"{type(exc).__name__}: {exc}"})
    with sqlite3.connect(CAS_DB) as cas_con:
        cas_substance_count = int(cas_con.execute("SELECT COUNT(*) FROM cas_substance").fetchone()[0])
    try:
        library = json.loads(RESULT_LIBRARY.read_text(encoding="utf-8"))
        program_known_substance_count = len(library.get("known_substances", {}))
    except Exception:
        program_known_substance_count = 0
    summary = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "model_db": str(MODEL_DB),
        "cas_db": str(CAS_DB),
        "model_count": len(model_rows),
        "completed_count": len(rows),
        "error_count": len(errors),
        "original_nonempty_count": sum(row["original_s2_status"] != "empty_result" for row in rows),
        "original_empty_count": sum(row["original_s2_status"] == "empty_result" for row in rows),
        "original_no_hazard_claim_count": sum(row["original_s2_no_hazard_claim"] for row in rows),
        "original_h_code_count": sum(bool(json.loads(row["original_s2_h_codes"])) for row in rows),
        "manual_direct_count": sum(row["legal_manual_status"] == "partial_direct_only" for row in rows),
        "manual_direct_no_flammable_count": sum(row["legal_manual_status"] == "partial_direct_no_flammable_only" for row in rows),
        "manual_insufficient_count": sum(row["legal_manual_status"] == "insufficient_legal_evidence" for row in rows),
        "cas_substance_count": cas_substance_count,
        "program_known_substance_count": program_known_substance_count,
        "program_status": dict(Counter(row["program_status"] for row in rows)),
        "mismatch_class": dict(Counter(row["mismatch_class"] for row in rows)),
        "word_source_status": dict(Counter(row["word_source_status"] for row in rows)),
        "direct_physical_mismatch_count": sum(
            row["legal_manual_status"] == "partial_direct_only"
            and {
                part.strip() for part in row["program_categories"].split("；") if "易燃液体" in part
            } != set(json.loads(row["legal_manual_categories"]))
            for row in rows
        ),
        "direct_pictogram_mismatch_count": sum(
            row["legal_manual_status"] == "partial_direct_only"
            and set(json.loads(row["legal_manual_pictograms"])) != set(json.loads(row["program_pictogram_codes"]))
            for row in rows
        ),
        "auto_pass_mismatch_count": sum(
            row["program_status"] == "auto_pass" and row["mismatch_class"] != "aligned_on_supported_facts"
            for row in rows
        ),
        "unsupported_h_code_frequency": dict(
            Counter(
                code
                for row in rows
                if row["mismatch_class"] == "program_overreach"
                for code in json.loads(row["program_h_codes"])
                if code not in set(json.loads(row["legal_manual_h_codes"]))
            )
        ),
        "legal_sources": LEGAL_SOURCES,
        "limitations": [
            "本轮只审计Section 2；不代表Section 3/9的内容正确性审计。",
            "原始结果来自正式数据库；Word路径存在时同时保留了直接读取状态。",
            "健康/环境分类需要补足物质级可追溯法规事实，当前不能仅凭CAS号或程序结果判定。",
        ],
    }
    _write_csv(out_dir / "section2_three_way_comparison.csv", rows)
    _write_csv(out_dir / "section2_errors.csv", errors)
    (out_dir / "section2_summary.json").write_text(_json(summary) + "\n", encoding="utf-8")
    (out_dir / "legal_source_register.json").write_text(_json(LEGAL_SOURCES) + "\n", encoding="utf-8")
    _write_xlsx(out_dir / "section2_three_way_comparison.xlsx", rows, summary)
    _write_report(out_dir / "section2_three_way_report.md", rows, summary)
    return summary


def main() -> int:
    parser = argparse.ArgumentParser(description="只读批量生成 Section 2 原始/法规人工/程序三方对照")
    parser.add_argument("--out-dir", type=Path, required=True, help="新的验证结果目录")
    args = parser.parse_args()
    summary = run(args.out_dir.expanduser().resolve())
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if summary["error_count"] == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
