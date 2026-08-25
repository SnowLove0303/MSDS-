"""Deterministic Section 2 inference functions.

This module is intentionally a rule layer, not a language model. It accepts
facts from ``s2_fact_package`` and returns values from ``s2_result_library``.
Every matched condition carries its inputs and legal basis so an unresolved
mixture threshold or missing source remains ``manual_review``.
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any, Iterable

from pictogram_registry import display_value, ordered_codes


ENGINE_DIR = Path(__file__).resolve().parent
LIBRARY_PATH = ENGINE_DIR / "s2_result_library.json"
LEGAL_SOURCE_DIR = ENGINE_DIR / "section2_legal_source"
LEGAL_RESULT_LIBRARY_PATH = LEGAL_SOURCE_DIR / "result_library.json"
LEGAL_RULE_CATALOG_PATH = LEGAL_SOURCE_DIR / "rule_catalog.json"
LEGAL_MANIFEST_PATH = LEGAL_SOURCE_DIR / "legal_source_manifest.json"
CAS_RE = re.compile(r"\b\d{2,7}-\d{2}-\d\b")
H_RE = re.compile(r"H\s*(\d{3})", re.IGNORECASE)
NUMBER_RE = re.compile(r"[-+]?\d+(?:[.,]\d+)?")


def _library() -> dict[str, Any]:
    """Load the structured legal source and retain legacy substance hints.

    The new legal source owns classifications, label mappings and rule metadata.
    The legacy library is intentionally retained only for the existing 12 CAS
    hints until those substances are migrated into the formal CAS result source.
    """
    legacy = json.loads(LIBRARY_PATH.read_text(encoding="utf-8"))
    legal = json.loads(LEGAL_RESULT_LIBRARY_PATH.read_text(encoding="utf-8"))
    rules = json.loads(LEGAL_RULE_CATALOG_PATH.read_text(encoding="utf-8"))
    manifest = json.loads(LEGAL_MANIFEST_PATH.read_text(encoding="utf-8"))
    merged = dict(legacy)
    merged.update(legal)
    merged["result_defaults"] = legal.get(
        "output_defaults", legacy.get("result_defaults", {})
    )
    merged["_legal_rules"] = rules.get("rules", [])
    merged["_legal_rule_catalog_version"] = rules.get("version", "unknown")
    merged["_legal_manifest"] = manifest
    return merged


def _legal_rule(library: dict[str, Any], rule_id: str) -> dict[str, Any]:
    for rule in library.get("_legal_rules", []):
        if rule.get("id") == rule_id:
            return rule
    return {"id": rule_id, "legal_sources": []}


def _rule_for_output(library: dict[str, Any], output_key: str) -> dict[str, Any]:
    for rule in library.get("_legal_rules", []):
        if rule.get("output_key") == output_key:
            return rule
    return {}


def _text(value: Any) -> str:
    if isinstance(value, list):
        return "；".join(_text(x) for x in value if _text(x))
    return str(value or "").strip()


def _first(facts: dict[str, Any], *names: str) -> str:
    for name in names:
        if name in facts:
            value = _text(facts[name])
            if value:
                return value
    return ""


def _dedupe(values: Iterable[str]) -> list[str]:
    output: list[str] = []
    for value in values:
        value = str(value or "").strip()
        if value and value not in output:
            output.append(value)
    return output


def extract_hazard_codes(value: Any) -> list[str]:
    return _dedupe(f"H{match.group(1)}" for match in H_RE.finditer(_text(value).upper()))


def _cas(value: Any) -> str:
    match = CAS_RE.search(_text(value))
    return match.group(0) if match else ""


def _number(value: Any) -> float | None:
    match = NUMBER_RE.search(_text(value).replace("，", ","))
    if not match:
        return None
    try:
        return float(match.group(0).replace(",", "."))
    except ValueError:
        return None


def _concentration_interval(value: Any) -> tuple[float, float] | None:
    text = _text(value).replace("％", "%").replace("±", "+/-")
    numbers = [float(x.replace(",", ".")) for x in NUMBER_RE.findall(text)]
    if not numbers:
        return None
    if "+/-" in text and len(numbers) >= 2:
        return max(0.0, numbers[0] - numbers[1]), numbers[0] + numbers[1]
    if re.search(r"\d\s*[-~至]\s*\d", text) and len(numbers) >= 2:
        return min(numbers[0], numbers[1]), max(numbers[0], numbers[1])
    return numbers[0], numbers[0]


def _temperature(value: Any) -> float | None:
    text = _text(value)
    if not text or text in {"无数据", "不适用", "无", "未知"}:
        return None
    # 不把“>94 ℃”当作 94 ℃，否则会错误判为类别 4。
    if re.search(r"[＞>]|[＜<]", text):
        return None
    return _number(text)


def _condition(rule_id: str, matched: bool, inputs: dict[str, Any], basis: list[str], result: Any = None) -> dict[str, Any]:
    row = {"id": rule_id, "matched": bool(matched), "inputs": inputs, "basis": basis}
    if result is not None:
        row["result"] = result
    return row


def _usage_rows(package: dict[str, Any]) -> list[dict[str, Any]]:
    return list((package.get("substance_snapshot") or {}).get("usage") or [])


def _decision_rows(package: dict[str, Any]) -> list[dict[str, Any]]:
    """读取可选的候选事实，不再从 CAS 身份库读取已移除的事实表。"""
    return list(package.get("decision_facts") or [])


def _known_substance_hints(package: dict[str, Any], library: dict[str, Any]) -> tuple[list[str], list[str], list[dict[str, Any]]]:
    codes: list[str] = []
    categories: list[str] = []
    conditions: list[dict[str, Any]] = []
    known = library.get("known_substances", {})
    for usage in _usage_rows(package):
        cas = _cas(usage.get("cas_no") or usage.get("raw_cas"))
        hint = known.get(cas)
        if not hint:
            conditions.append(_condition(
                "known_substance_lookup",
                False,
                {"cas": cas or "无数据", "raw_name": usage.get("raw_name") or "无数据"},
                ["s2_result_library.known_substances"],
            ))
            continue
        hint_codes = [str(x) for x in hint.get("hazard_codes", [])]
        hint_categories = [str(x) for x in hint.get("categories", [])]
        codes.extend(hint_codes)
        categories.extend(hint_categories)
        conditions.append(_condition(
            "known_substance_lookup",
            True,
            {
                "cas": cas,
                "raw_name": usage.get("raw_name") or "无数据",
                "concentration": usage.get("concentration") or "无数据",
            },
            ["s2_result_library.known_substances", "hazards_lib.md"],
            {"hazard_codes": hint_codes, "categories": hint_categories},
        ))
    return _dedupe(codes), _dedupe(categories), conditions


def _flammable_liquid_rule(package: dict[str, Any], library: dict[str, Any]) -> tuple[str, str, dict[str, Any]]:
    s9 = ((package.get("input_snapshot") or {}).get("s9") or {}).get("facts") or {}
    flash_raw = _first(s9, "闪点")
    boil_raw = _first(s9, "初沸点")
    flash = _temperature(flash_raw)
    boil = _temperature(boil_raw)
    output_key = ""
    if flash is not None:
        if flash < 23 and boil is not None:
            output_key = "flammable_liquid_category_1" if boil <= 35 else "flammable_liquid_category_2"
        elif 23 <= flash <= 60:
            output_key = "flammable_liquid_category_3"
        elif 60 < flash <= 93:
            output_key = "flammable_liquid_category_4"
        elif flash > 93:
            output_key = "not_classified_under_endpoint"
    result_library = library.get("classification_results", {})
    output = result_library.get(output_key, {}) if output_key else {}
    category = str(output.get("label") or "")
    h_code = str(output.get("h_code") or "")
    if output_key == "not_classified_under_endpoint":
        # 该结果是某一危害端点的否定，不是 Section 2 的危险类别文本。
        category = ""
    if output_key:
        rule = _rule_for_output(library, output_key)
        rule_id = str(rule.get("id") or "flammable_liquid_by_s9")
        result = {"output_key": output_key, **output}
        basis = [str(source.get("id")) for source in rule.get("legal_sources", [])]
        basis.append("input_snapshot.s9")
    else:
        rule = _legal_rule(library, "FL-005")
        rule_id = str(rule.get("id") or "FL-005")
        result = "manual_review"
        basis = [str(source.get("id")) for source in rule.get("legal_sources", [])]
        basis.append("input_snapshot.s9")
    condition = _condition(
        rule_id,
        bool(output_key),
        {"闪点": flash_raw or "无数据", "初沸点": boil_raw or "无数据", "闪点数值": flash, "初沸点数值": boil},
        basis,
        result,
    )
    return category, h_code, condition


def _categories_from_codes(codes: Iterable[str]) -> list[str]:
    mapping = {
        "H224": "易燃液体，类别 1", "H225": "易燃液体，类别 2", "H226": "易燃液体，类别 3", "H227": "易燃液体，类别 4",
        "H300": "急性毒性，类别 1", "H301": "急性毒性，类别 3", "H302": "急性毒性，类别 4",
        "H303": "急性毒性，类别 5", "H304": "吸入危害，类别 1", "H305": "吸入危害，类别 2",
        "H310": "急性毒性，类别 1", "H311": "急性毒性，类别 3", "H312": "急性毒性，类别 4",
        "H314": "皮肤腐蚀，类别 1", "H315": "皮肤刺激，类别 2", "H316": "皮肤刺激，类别 3",
        "H317": "皮肤致敏，类别 1", "H318": "严重眼损伤，类别 1", "H319": "严重眼刺激，类别 2",
        "H320": "眼刺激，类别 2", "H330": "急性毒性，类别 1", "H331": "急性毒性，类别 3",
        "H332": "急性毒性，类别 4", "H335": "特异性靶器官毒性—一次接触，类别 3",
        "H336": "特异性靶器官毒性—一次接触，类别 3", "H340": "生殖细胞致突变性，类别 1",
        "H341": "生殖细胞致突变性，类别 2", "H350": "致癌性，类别 1", "H351": "致癌性，类别 2",
        "H360": "生殖毒性，类别 1", "H361": "生殖毒性，类别 2", "H370": "特异性靶器官毒性—一次接触，类别 1",
        "H371": "特异性靶器官毒性—一次接触，类别 2", "H372": "特异性靶器官毒性—反复接触，类别 1",
        "H373": "特异性靶器官毒性—反复接触，类别 2", "H400": "对水环境的危害，急性 1",
        "H401": "对水环境的危害，急性 2", "H402": "对水环境的危害，急性 3",
        "H410": "对水环境的危害，慢性 1", "H411": "对水环境的危害，慢性 2",
        "H412": "对水环境的危害，慢性 3", "H413": "对水环境的危害，慢性 4",
    }
    return _dedupe(mapping[code] for code in codes if code in mapping)


def _pictograms_for_codes(codes: Iterable[str], library: dict[str, Any]) -> list[str]:
    output: list[str] = []
    source_mapping = library.get("h_code_to_pictograms", {})
    for code in codes:
        if code in source_mapping:
            output.extend(str(value) for value in source_mapping[code])
            continue
        number = int(code[1:])
        if code in {"H224", "H225", "H226"} or 200 <= number <= 226:
            output.append("GHS02")
        elif 270 <= number <= 272:
            output.append("GHS03")
        elif 280 <= number <= 281:
            output.append("GHS04")
        elif code in {"H314", "H318"}:
            output.append("GHS05")
        elif code in {"H300", "H301", "H310", "H311", "H330", "H331"}:
            output.append("GHS06")
        elif code in {"H302", "H303", "H312", "H313", "H315", "H316", "H317", "H319", "H320", "H332", "H335", "H336"}:
            output.append("GHS07")
        elif code in {"H304", "H305", "H340", "H341", "H350", "H351", "H360", "H361", "H370", "H371", "H372", "H373"}:
            output.append("GHS08")
        elif 400 <= number <= 413:
            output.append("GHS09")
    return ordered_codes(output)


def _signal_word(codes: set[str], categories: list[str], library: dict[str, Any]) -> str:
    if codes & {"H224", "H225", "H300", "H301", "H310", "H311", "H330", "H331", "H314", "H318", "H340", "H350", "H360", "H370", "H372"}:
        return library["signal_words"]["danger"]
    if codes or any("类别 3" in x or "类别 4" in x for x in categories):
        return library["signal_words"]["warning"]
    return library["signal_words"]["none"]


def _precaution_codes(codes: set[str], pictograms: list[str]) -> list[str]:
    output: list[str] = []
    if "GHS02" in pictograms:
        output.append("P210")
    if codes & {"H302", "H312", "H313", "H315", "H316", "H317", "H318", "H319", "H320", "H332", "H335", "H336", "H360", "H361", "H370", "H371", "H372", "H373"}:
        output.append("P280")
    if codes & {"H304", "H305", "H330", "H331", "H332", "H335", "H336"}:
        output.append("P261")
    if codes & {"H315", "H316", "H317", "H312", "H313"}:
        output.append("P302+P352")
    if codes & {"H314", "H318", "H319", "H320"}:
        output.append("P305+P351+P338")
    if codes & {"H330", "H331", "H332"}:
        output.append("P304+P340")
    if "GHS09" in pictograms:
        output.extend(["P273", "P391"])
    if codes:
        output.append("P501")
    return _dedupe(output)


def _mixture_review(package: dict[str, Any], codes: list[str], library: dict[str, Any]) -> tuple[list[dict[str, Any]], list[str]]:
    """Apply the GB/T 17519 threshold gate without silently over-classifying."""
    conditions: list[dict[str, Any]] = []
    review: list[str] = []
    usages = _usage_rows(package)
    thresholds = library.get("mixture_thresholds_by_h_code", {})
    known = library.get("known_substances", {})
    for usage in usages:
        cas = _cas(usage.get("cas_no") or usage.get("raw_cas"))
        usage_codes = [str(x) for x in (known.get(cas) or {}).get("hazard_codes", [])]
        concentration = _concentration_interval(usage.get("concentration"))
        if not usage_codes:
            continue
        if not concentration:
            if cas in {""}:
                continue
            review.append(f"CAS {cas} 缺少有效含量，不能完成混合物阈值判定。")
            conditions.append(_condition(
                "mixture_threshold",
                False,
                {"cas": cas, "concentration": usage.get("concentration") or "无数据"},
                ["GB/T 17519-2013"],
                "需要含量",
            ))
            continue
        low, high = concentration
        for code in usage_codes:
            threshold = thresholds.get(code)
            if threshold is None:
                continue
            if high < threshold:
                matched = False
                state = "低于阈值"
            elif low >= threshold:
                matched = True
                state = "达到阈值"
            else:
                matched = True
                state = "区间跨越阈值"
                review.append(f"CAS {cas} 的 {code} 含量区间跨越 {threshold:g}% 阈值。")
            conditions.append(_condition(
                "mixture_threshold",
                matched,
                {"cas": cas, "concentration_low": low, "concentration_high": high, "code": code, "threshold_percent": threshold},
                ["GB/T 17519-2013"],
                state,
            ))
    return conditions, _dedupe(review)


def infer_s2(package: dict[str, Any]) -> dict[str, Any]:
    """Evaluate the current fact package and return normalized Section 2 values."""
    library = _library()
    model = str(package.get("model") or "")
    known_codes, known_categories, conditions = _known_substance_hints(package, library)
    candidate_codes = _dedupe(
        str(fact.get("result_code") or "").upper()
        for fact in _decision_rows(package)
        if re.fullmatch(r"H\d{3}", str(fact.get("result_code") or "").upper())
    )
    flammable_category, flash_code, flash_condition = _flammable_liquid_rule(package, library)
    conditions.append(flash_condition)
    codes = _dedupe([*known_codes, flash_code])
    categories = _dedupe([*known_categories, flammable_category, *_categories_from_codes(codes)])
    pictograms = _pictograms_for_codes(codes, library)
    threshold_conditions, threshold_review = _mixture_review(package, codes, library)
    conditions.extend(threshold_conditions)

    hazard_values = library.get("hazard_statements", {})
    precaution_values = library.get("precautionary_statements", {})
    ordered_hazard_codes = sorted(codes, key=lambda code: int(code[1:]))
    precaution_codes = _precaution_codes(set(codes), pictograms)
    hazard_statements = [hazard_values[code] for code in ordered_hazard_codes if code in hazard_values]
    precaution_statements = [precaution_values[code] for code in precaution_codes if code in precaution_values]
    signal_word = _signal_word(set(codes), categories, library)

    physical_categories = [x for x in categories if "易燃" in x or "爆炸" in x or "氧化" in x or "加压" in x]
    health_categories = [x for x in categories if not ("水环境" in x or "易燃" in x or "爆炸" in x or "氧化" in x or "加压" in x)]
    environment_categories = [x for x in categories if "水环境" in x]
    pictogram_value = (
        display_value(pictograms)
        if pictograms
        else (library["result_defaults"]["no_hazard"] if codes else library["result_defaults"]["no_data"])
    )
    result = {
        "GHS危险性类别": "；".join(categories) or library["result_defaults"]["no_data"],
        "象形图": pictogram_value,
        "信号词": signal_word if codes else library["signal_words"]["unknown"],
        "危险性说明": "\n".join(hazard_statements) or library["result_defaults"]["no_data"],
        "防范说明": "\n".join(precaution_statements) or library["result_defaults"]["no_data"],
        "物理和化学危险": "；".join(physical_categories) + "。" if physical_categories else library["result_defaults"]["no_data"],
        "健康危害": "；".join(health_categories) + "。" if health_categories else library["result_defaults"]["no_hazard"],
        "环境危害": "；".join(environment_categories) + "。" if environment_categories else library["result_defaults"]["no_hazard"],
        "其他危害": library["result_defaults"]["no_data"],
    }
    review_reasons = list(threshold_review)
    if any(
        c.get("id") == "cas_structured_hazard_candidate"
        and c.get("inputs", {}).get("validation_status") != "validated"
        for c in conditions
    ):
        review_reasons.append("CAS 危害分类来自外部候选事实，尚未完成中国适用标准核验。")
    if not codes:
        review_reasons.append("S1/S3/S9/CAS 事实中未形成可执行的 H 码或物理危险判据。")
    if any(c["id"] == "known_substance_lookup" and not c["matched"] for c in conditions):
        review_reasons.append("存在未命中结果库的 CAS 或保密组分，需补充物质危害事实。")
    if codes and any("区间跨越" in str(c.get("result")) for c in conditions if c["id"] == "mixture_threshold"):
        review_reasons.append("混合物含量区间跨越法规阈值，不能自动确定最终分类。")
    status = "manual_review" if review_reasons else "auto_pass"
    if not codes:
        status = "insufficient_evidence"
    return {
        "model": model,
        "rule_version": library.get("_legal_rule_catalog_version", "unknown"),
        "status": status,
        "conditions": conditions,
        "result": result,
        "candidate_label_elements": {
            "pictogram_codes": pictograms,
            "reference_candidate_hazard_codes": candidate_codes,
            "pictograms": result["象形图"],
            "signal_word": result["信号词"],
            "hazard_statements": hazard_statements,
            "precautionary_statements": precaution_statements,
        },
        "review_reasons": _dedupe(review_reasons),
        "legal_basis": {
            "manifest_version": (library.get("_legal_manifest") or {}).get("schema_version", "unknown"),
            "source_ids": [source.get("id") for source in (library.get("_legal_manifest") or {}).get("legal_sources", [])],
            "rule_catalog_version": library.get("_legal_rule_catalog_version", "unknown"),
            "legacy_reference": library.get("rule_basis", {}),
        },
    }


__all__ = ["extract_hazard_codes", "infer_s2"]
