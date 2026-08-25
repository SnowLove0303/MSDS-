"""Read-only validator for the structured Section 2 legal source."""

from __future__ import annotations

import json
import operator
from pathlib import Path
from typing import Any, Callable


ROOT = Path(__file__).resolve().parent


def load(name: str) -> dict[str, Any]:
    return json.loads((ROOT / name).read_text(encoding="utf-8"))


def compare(value: float, op: str, boundary: float) -> bool:
    operations: dict[str, Callable[[float, float], bool]] = {
        "<": operator.lt,
        "<=": operator.le,
        ">": operator.gt,
        ">=": operator.ge,
        "=": operator.eq,
    }
    return operations[op](value, boundary)


def classify_flammable(facts: dict[str, Any], result_library: dict[str, Any]) -> dict[str, Any]:
    flash = facts.get("flash_point") or {}
    value = flash.get("value")
    if not isinstance(value, (int, float)) or flash.get("status") == "unknown":
        return {"status": "manual_review", "reason": "闪点缺失或不可比较"}

    if value < 23:
        boil = facts.get("initial_boiling_point") or {}
        boil_value = boil.get("value")
        if not isinstance(boil_value, (int, float)):
            return {"status": "manual_review", "reason": "闪点低于23℃但初沸点缺失"}
        output_key = "flammable_liquid_category_1" if boil_value <= 35 else "flammable_liquid_category_2"
    elif value <= 60:
        output_key = "flammable_liquid_category_3"
    elif value <= 93:
        output_key = "flammable_liquid_category_4"
    else:
        output_key = "not_classified_under_endpoint"

    result = result_library["classification_results"][output_key]
    return {"status": "auto_pass", "output_key": output_key, **result}


def assert_equal(actual: Any, expected: Any, message: str) -> None:
    if actual != expected:
        raise AssertionError(f"{message}: expected={expected!r}, actual={actual!r}")


def validate_structure() -> dict[str, int]:
    manifest = load("legal_source_manifest.json")
    facts = load("fact_schema.json")
    rules = load("rule_catalog.json")
    results = load("result_library.json")
    policies = load("evidence_policies.json")
    tests = load("test_cases.json")

    assert_equal(manifest["scope"], "Section 2 危险性概述", "法规源范围")
    assert_equal(facts["fact_policy"]["unknown_is_not_no_hazard"], True, "未知数据政策")
    assert_equal(policies["release_gate"][1].startswith("象形图"), True, "发布门禁")
    assert_equal(results["classification_results"]["flammable_liquid_category_4"]["h_code"], "H227", "类别4 H代码")
    assert_equal(results["classification_results"]["flammable_liquid_category_4"]["pictograms"], [], "类别4象形图")
    assert_equal(results["h_code_to_pictograms"]["H227"], [], "H227象形图映射")
    assert_equal(set(results["hazard_statements"]), set(results["h_code_to_pictograms"]), "H代码映射完整性")

    rule_ids = {rule["id"] for rule in rules["rules"]}
    assert_equal(len(rule_ids), len(rules["rules"]), "规则编号唯一性")
    source_ids = {source["id"] for source in manifest["legal_sources"]}
    for rule in rules["rules"]:
        for source in rule.get("legal_sources", []):
            if source["id"] not in source_ids:
                raise AssertionError(f"规则 {rule['id']} 引用了未登记法规源 {source['id']}")

    passed = 0
    for case in tests["cases"]:
        result = classify_flammable(case["facts"], results)
        expected = case["expected"]
        for key in ("status", "output_key", "h_code", "signal_word", "pictograms"):
            if key in expected:
                assert_equal(result.get(key), expected[key], f"案例 {case['id']} 字段 {key}")
        if expected.get("inherit_component_h_code") is False:
            mix_rule = next(rule for rule in rules["rules"] if rule["id"] == "MIX-004")
            assert_equal(mix_rule["decision"], "do_not_inherit_component_h_code", f"案例 {case['id']} 组分继承禁止")
        passed += 1

    return {
        "legal_sources": len(manifest["legal_sources"]),
        "rules": len(rules["rules"]),
        "h_codes": len(results["hazard_statements"]),
        "test_cases_passed": passed,
        "test_cases_total": len(tests["cases"]),
    }


if __name__ == "__main__":
    summary = validate_structure()
    print(json.dumps({"status": "passed", **summary}, ensure_ascii=False, indent=2))
