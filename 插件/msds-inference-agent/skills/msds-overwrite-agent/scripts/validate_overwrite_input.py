#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Validate the Agent-facing source/overwrite contract without writing a document."""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

S9_FIELDS = {
    "外观", "嗅觉阈值", "pH值", "离子性", "初沸点", "闪点", "蒸发速率",
    "可燃性（固态、气态）", "燃烧值", "饱和蒸气压", "相对蒸气密度", "密度",
    "水溶性", "表面张力", "辛醇/水分配系数对数值", "自燃温度", "引燃温度",
    "分解温度", "动力粘度", "爆炸特性", "粉尘爆炸级别", "固体含量", "有效成分",
    "玻璃化温度", "最低成膜温度", "NCO含量", "羟基含量", "熔点/凝固点", "酸值",
    "碘值", "倾点", "浊点", "分子量", "含量", "APHA值", "HLB值", "其他信息",
}
PLACEHOLDERS = {"无数据", "无数据资料", "无可用数据", "无资料", "未提供", "不适用",
                "no data", "no data available", "not applicable", "not available"}


def clean(value) -> str:
    return str(value or "").strip()


def is_placeholder(value) -> bool:
    return clean(value).casefold() in {x.casefold() for x in PLACEHOLDERS}


def validate(payload: dict) -> list[str]:
    errors: list[str] = []
    if not isinstance(payload, dict):
        return ["payload 必须是 JSON 对象"]
    sections = payload.get("sections")
    if not isinstance(sections, dict):
        return ["sections 必须是对象"]
    for sec in ("1", "3", "9"):
        if sec not in sections:
            errors.append(f"缺少入口 Section {sec}")
    s1 = sections.get("1")
    if not isinstance(s1, list):
        errors.append("Section 1 必须是字段列表")
    s3 = sections.get("3")
    if not isinstance(s3, dict):
        errors.append("Section 3 必须是对象")
    elif not isinstance(s3.get("components", []), list):
        errors.append("Section 3.components 必须是列表")
    s9 = sections.get("9")
    if not isinstance(s9, list):
        errors.append("Section 9 必须是字段列表")
    else:
        for index, item in enumerate(s9, 1):
            if not isinstance(item, dict):
                errors.append(f"Section 9 第 {index} 项必须是对象")
                continue
            label = clean(item.get("label"))
            if label not in S9_FIELDS:
                errors.append(f"Section 9 字段不属于数据库 37 项: {label}")
            if is_placeholder(item.get("value")):
                errors.append(f"Section 9 {label} 不能把占位词作为正式值")
    for sec in ("11", "12"):
        for item in sections.get(sec, []) if isinstance(sections.get(sec), list) else []:
            if isinstance(item, dict) and is_placeholder(item.get("value")):
                # 缺失研究应由引擎生成控制句，不接受占位词进入正式 payload。
                errors.append(f"Section {sec} {item.get('label', '')} 不能使用缺失占位词")
    choices = payload.get("output", payload.get("product", {}))
    if isinstance(choices, dict):
        if choices.get("brand") and choices["brand"] not in {"guanzhi", "guocai"}:
            errors.append("brand 只能是 guanzhi 或 guocai")
        if choices.get("language") and choices["language"] not in {"zh", "en"}:
            errors.append("language 只能是 zh 或 en")
        if choices.get("output_format") and choices["output_format"] not in {"docx", "pdf"}:
            errors.append("output_format 只能是 docx 或 pdf")
    return errors


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--payload", required=True, type=Path)
    args = parser.parse_args()
    payload = json.loads(args.payload.read_text(encoding="utf-8"))
    errors = validate(payload)
    if errors:
        print(json.dumps({"ok": False, "errors": errors}, ensure_ascii=False, indent=2))
        return 1
    print(json.dumps({"ok": True, "message": "Agent input contract passed"}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
