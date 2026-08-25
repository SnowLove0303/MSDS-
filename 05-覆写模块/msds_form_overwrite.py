#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""17 节 MSDS 控制台表单 → write_items → Word 模板覆写。

该脚本把现有结构读取器、17 节骨架和覆写引擎串成一条可交互流水线：

1. 读取指定模板，按节/父级/标签生成可编辑字段目录；
2. 通过中文控制台逐项录入或保留模板值；
3. 生成覆写引擎契约 write_items JSON；
4. 用户确认后调用现有覆写引擎；
5. 通过结构回读校验输出 Word。

默认不把空白输入写入文档，按 Enter 表示保留模板原值；输入 ``!清空``
可显式清空字段。默认模板是项目确认的唯一中文冠志模板。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

BASE_DIR = Path(os.environ.get(
    "MSDS_ROOT", str(Path(__file__).resolve().parents[1])))
READ_DIR = BASE_DIR / "02-检索系统"
ENGINE_DIR = BASE_DIR / "05-覆写模块"
DEFAULT_TEMPLATE = BASE_DIR / "03-数据库" / "正式库" / "推导方案" / "PEA-4139 MSDS_CN 冠志 模板.docx"
DEFAULT_KEEP_STRUCTURE = [1, 2, 3, 4, 5, 6, 7, 8, 10, 11, 12, 13, 14, 15, 16]
BIO_HEADER = ["组分名称", "标准来源", "生物监测指标", "生物限值", "采样时间"]

for _path in (READ_DIR, ENGINE_DIR):
    if str(_path) not in sys.path:
        sys.path.insert(0, str(_path))

from core.docx_reader import read_msds  # noqa: E402
from core.schema import standard_fields, standard_name  # noqa: E402
from core.edit_policy import (  # noqa: E402
    is_editable_note,
    is_editable_field,
    is_formal_placeholder,
    validate_write_items_permissions,
)
from core.msds_db import listed_rows_from_result  # noqa: E402
import msds_overwrite_engine as overwrite_engine  # noqa: E402


@dataclass
class EditableField:
    """一个可交互字段及其模板上下文。"""

    section: int
    seq: str
    label: str
    parent: str
    current: str = ""
    kind: str = "field"
    key: str = ""

    @property
    def display(self) -> str:
        """返回用于控制台的父子级显示文本。"""
        prefix = f"{self.seq} " if self.seq else ""
        parent = f"{self.parent} > " if self.parent else ""
        return f"S{self.section} | {parent}{prefix}{self.label}"


@dataclass
class FormState:
    """交互表单的暂存状态。"""

    template: Path
    result: Any
    fields: dict[str, EditableField] = field(default_factory=dict)
    values: dict[str, str] = field(default_factory=dict)
    touched_sections: set[int] = field(default_factory=set)
    cleared_keys: set[str] = field(default_factory=set)
    components: list[dict[str, str]] = field(default_factory=list)
    product_type: str = "混合物"
    bio_rows: list[list[str]] = field(default_factory=list)
    laws: list[str] = field(default_factory=list)
    s15_special: dict[str, str] = field(default_factory=dict)
    # S16 的免责声明是模板 note 槽位，但必须作为普通输入字段暴露给
    # Python GUI、Web GUI 和 Agent payload；这里保留显式别名便于接口审计。
    s16_special: dict[str, str] = field(default_factory=dict)
    # S9 源文件/用户输入的特殊标签表述；输出仍按固定模板标准标签匹配。
    label_overrides: dict[str, str] = field(default_factory=dict)


def _clean(text: Any) -> str:
    """规范化控制台文本，但保留用户输入的内部换行。"""
    return str(text or "").replace("\r\n", "\n").replace("\r", "\n").strip()


def _parse_sections(raw: str | None) -> list[int]:
    """解析节选择，支持 ``all``、``0-16`` 和逗号列表。"""
    if not raw or raw.lower() in {"all", "全部", "0-16", "0~16"}:
        return list(range(17))
    out: set[int] = set()
    for part in re.split(r"[,，\s]+", raw):
        if not part:
            continue
        if "-" in part or "~" in part:
            a, b = re.split(r"[-~]", part, maxsplit=1)
            out.update(range(int(a), int(b) + 1))
        else:
            out.add(int(part))
    invalid = sorted(x for x in out if x < 0 or x > 16)
    if invalid:
        raise ValueError(f"节号必须在 0~16 范围内: {invalid}")
    return sorted(out)


def _field_key(section: int, seq: str, label: str, index: int) -> str:
    """生成同标签字段的稳定交互键。"""
    base = f"S{section}|{seq}|{label}".strip("|")
    return base if index == 0 else f"{base}#{index + 1}"


def _template_rows(result: Any, section: int) -> list[Any]:
    """按冻结 17 节骨架获取模板行，并适配批准英文模板的 S1 标签。

    检索系统的 17 节骨架仍以中文标准字段为内部契约；批准的英文模板
    在 S1.1 将原 ``Chinese name`` 子字段调整为 ``Product name``。这里只
    在覆写表单层做局部显示/写入标签适配，不修改检索系统的全局冻结骨架。
    """
    rows = listed_rows_from_result(result, section, s9_active_only=False, active_only=False)
    if section == 1 and _is_adjusted_english_section1(result):
        for row in rows:
            if row.kind == "field" and row.label == "中文名称":
                row.label = "Product name"
                break
    return rows


def _is_adjusted_english_section1(result: Any) -> bool:
    """识别当前批准英文模板的 S1.1 Product name 结构。"""
    section = getattr(result, "sections", {}).get(1)
    if section is None:
        return False
    labels = {_clean(field.label).casefold() for field in getattr(section, "fields", [])}
    return "product name" in labels and "chinese name" not in labels


def _load_state(template: Path) -> FormState:
    """读取模板并构造字段、成分和特殊子表的初始状态。"""
    result = read_msds(template)
    state = FormState(template=template, result=result)
    for section in range(17):
        parent = ""
        seen: dict[tuple[str, str], int] = {}
        for row in _template_rows(result, section):
            if row.kind == "sub":
                parent = f"{row.seq} {row.label}".strip()
                continue
            if row.kind == "subtable":
                if section == 8 and row.label == "生物限值":
                    state.bio_rows = [list(x) for x in row.sub_rows if any(_clean(v) for v in x)]
                continue
            if row.kind not in {"field", "note"}:
                continue
            label = _clean(row.label)
            if section == 0 and label == "页码":
                continue
            if not label:
                label = "通栏说明"
            if section == 15 and row.kind == "note":
                # S15 的结构行不走普通字段白名单。法规条目和两个可编辑
                # 说明槽位由专用容器管理，避免条目被普通 note 逻辑丢弃。
                if label.startswith("法规条目"):
                    value = _clean(row.value)
                    if value and not is_formal_placeholder(value):
                        state.laws.append(value)
                    continue
                if label in {"其它的规定", "符合下列法规要求"}:
                    value = _clean(row.value)
                    if value == label:
                        value = ""
                    state.s15_special[label] = value
                    continue
            if not (
                is_editable_note(section, label, _clean(row.value))
                if row.kind == "note"
                else is_editable_field(section, label, kind=row.kind)
            ):
                continue
            pair = (row.seq or "", label)
            occurrence = seen.get(pair, 0)
            seen[pair] = occurrence + 1
            key = _field_key(section, pair[0], pair[1], occurrence)
            field_item = EditableField(
                section=section,
                seq=_clean(row.seq),
                label=label,
                parent=parent,
                current=_clean(row.value),
                kind=row.kind,
                key=key,
            )
            state.fields[key] = field_item
            state.values[key] = field_item.current
            if section == 16 and row.kind == "note":
                state.s16_special[label] = field_item.current

    # Web/CLI 与 Python GUI 共用同一份数据库标准 S9 目录。批准模板可以
    # 为版式省略 9.23~9.37 的物理行，但表单契约不能因此丢失 37 项；
    # 缺少的字段作为空值输入槽位，实际有值时由引擎按模板参考行插入。
    existing_s9 = {
        (standard_name(9, item.label) or item.label)
        for item in state.fields.values() if item.section == 9
    }
    for index, schema_field in enumerate(standard_fields(9), start=1):
        label = _clean(schema_field.name)
        if label in existing_s9:
            continue
        key = f"S9|9.{index}|{label}"
        state.fields[key] = EditableField(
            section=9, seq=f"9.{index}", label=label, parent="",
            current="", kind="field", key=key,
        )
        state.values[key] = ""
        existing_s9.add(label)

    sec3 = result.section(3)
    if sec3 is not None:
        for row in sec3.iter_rows():
            if row.kind == "field" and _clean(row.label) in {"产品类型", "产品类型："}:
                state.product_type = _clean(row.value) or "混合物"
        for comp in sec3.components:
            state.components.append({
                "name": _clean(comp.name),
                "cas": _clean(comp.cas),
                "conc": _clean(comp.conc),
            })
    return state


def _input(prompt: str, default: str = "") -> str:
    """读取一行中文控制台输入，默认值仅用于提示。"""
    suffix = f" [{default}]" if default else ""
    return input(f"{prompt}{suffix}: ").strip()


def _multiline(prompt: str, default: str = "") -> str:
    """读取多行字段；空行结束，直接 Enter 保留默认值。"""
    print(f"{prompt}（多行输入，单独输入【结束】完成；直接 Enter 保留当前值）")
    if default:
        print(f"当前值：{default}")
    first = input("> ")
    if not first:
        return default
    lines = [] if first == "!清空" else [first]
    while True:
        line = input("> ")
        if line == "结束":
            break
        lines.append(line)
    return "\n".join(lines).strip()


def _edit_generic_section(state: FormState, section: int) -> None:
    """交互编辑一个普通节的 field/note 字段。"""
    fields = [x for x in state.fields.values() if x.section == section]
    if not fields:
        print(f"S{section} 没有可交互字段。")
        return
    print(f"\n===== S{section} 字段录入 =====")
    changed = False
    for item in fields:
        print(f"\n{item.display}")
        value = _multiline("输入内容", state.values[item.key])
        if value == "!清空":
            value = ""
            state.cleared_keys.add(item.key)
            changed = True
        else:
            state.cleared_keys.discard(item.key)
            if value != state.values[item.key]:
                changed = True
        state.values[item.key] = value
    if changed:
        state.touched_sections.add(section)


def _edit_s3(state: FormState) -> None:
    """交互编辑 S3 产品类型和成分子表。"""
    print("\n===== S3 成分/组成资料 =====")
    product_type = _input("产品类型（混合物/单质/化合物/未知）", state.product_type)
    if product_type:
        state.product_type = product_type
    old = state.components
    state.components = []
    print("逐行输入：化学品名称 | CAS编号 | 含量；名称为空结束。")
    index = 0
    while True:
        old_row = old[index] if index < len(old) else {"name": "", "cas": "", "conc": ""}
        name = _input(f"成分 {index + 1} 名称", old_row["name"])
        if not name:
            break
        cas = _input("CAS 编号", old_row["cas"])
        conc = _input("含量（% w/w）", old_row["conc"])
        state.components.append({"name": name, "cas": cas, "conc": conc})
        index += 1
    state.touched_sections.add(3)


def _edit_s8_bio(state: FormState) -> bool:
    """编辑 S8.2 生物限值子表。"""
    print("\nS8.2 生物限值：每行按 组分|标准来源|监测指标|限值|采样时间 输入；空行结束。")
    old = state.bio_rows
    rows: list[list[str]] = []
    index = 0
    while True:
        default = "|".join(old[index]) if index < len(old) else ""
        raw = _input(f"生物限值 {index + 1}", default)
        if not raw:
            break
        parts = [x.strip() for x in raw.split("|")]
        if len(parts) != 5:
            print("必须正好输入 5 列，使用 | 分隔。")
            continue
        rows.append(parts)
        index += 1
    changed = rows != old
    state.bio_rows = rows
    return changed


def _edit_s15(state: FormState) -> None:
    """编辑 S15 法规条目；指引段和结构标题按规范锁定。"""
    print("\n===== S15 法规信息 =====")
    laws: list[str] = []
    print("逐条输入法规条目；空行结束。")
    while True:
        value = _input(f"法规条目 {len(laws) + 1}")
        if not value:
            break
        laws.append(value)
    state.s15_special = {}
    state.laws = laws
    state.touched_sections.add(15)


def _edit_selected_sections(state: FormState, sections: list[int]) -> None:
    """按用户选择编辑全部指定节。"""
    for section in sections:
        if section == 3:
            _edit_s3(state)
        elif section == 8:
            _edit_generic_section(state, section)
            if _edit_s8_bio(state):
                state.touched_sections.add(8)
        elif section == 15:
            _edit_s15(state)
        else:
            _edit_generic_section(state, section)


def _items_for_section(state: FormState, section: int) -> list[dict[str, str]]:
    """把普通字段转为覆写引擎平铺项；正式输出不写缺失占位词。"""
    items: list[dict[str, str]] = []
    for item in state.fields.values():
        if item.section != section:
            continue
        value = _clean(state.values[item.key])
        current = _clean(item.current)
        if item.key not in state.cleared_keys and value == current and value and section != 9:
            # 普通节未改动的模板值无需再次写入；S9 例外，因正式输出
            # 物理删除空行，必须把现有真实值一并带入本次标准 37 项提交。
            continue
        if not value or is_formal_placeholder(value):
            # 空值只表示该字段不进入正式呈现。S11/S12 在节级别由
            # build_write_items 生成唯一的证据缺失兜底行。
            continue
        label = state.label_overrides.get(item.key, item.label) if section == 9 else item.label
        items.append({"seq": item.seq, "label": label, "value": value})
    return items


def build_write_items(state: FormState) -> dict[str, Any]:
    """从交互状态生成标准 write_items 契约。"""
    sections: dict[str, Any] = {}
    for section in sorted(state.touched_sections):
        if section == 3:
            sections["3"] = {
                "产品类型": state.product_type or "混合物",
                "components": state.components,
            }
            continue
        if section == 8:
            items = _items_for_section(state, 8)
            if state.bio_rows:
                items.append({
                    "seq": "8.2",
                    "label": "生物限值",
                    "value": "",
                    "subtable": {"header": BIO_HEADER, "rows": state.bio_rows},
                })
            sections["8"] = items
            continue
        if section == 15:
            items = _items_for_section(state, 15)
            for label, value in state.s15_special.items():
                if not is_formal_placeholder(value):
                    items.append({"seq": "", "label": label, "value": value})
            for law in state.laws:
                if law and not is_formal_placeholder(law):
                    items.append({"seq": "", "label": "法规条目", "value": law})
            sections["15"] = items
            continue
        items = _items_for_section(state, section)
        if section in {11, 12} and not items:
            items = [{
                "seq": "",
                "label": "产品说明",
                "value": ("该产品无可用的毒理学研究。"
                           if section == 11 else "该产品无可用的生态毒理学研究。"),
            }]
        if items:
            sections[str(section)] = items

    return {
        "sections": sections,
        "keep_structure": DEFAULT_KEEP_STRUCTURE,
        "empty_policy": "overwrite" if state.cleared_keys else "preserve",
        "missing_policy": "preserve",
        "missing_text": "",
        "formal_output_policy": {
            "suppress_placeholders": True,
            "show_s9_only_with_value": True,
            "evidence_fallback": True,
        },
    }


def _load_input_json(path: Path, state: FormState) -> dict[str, Any]:
    """读取预先准备的 write_items；用于自动化和重复产出。"""
    payload = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict) or not isinstance(payload.get("sections"), dict):
        raise ValueError("输入 JSON 必须是 write_items 对象，并包含 sections")
    return payload


def _preview(payload: dict[str, Any]) -> None:
    """输出覆写前的摘要。"""
    print("\n===== 覆写确认摘要 =====")
    sections = payload.get("sections", {})
    for sec, items in sections.items():
        if isinstance(items, dict):
            comps = items.get("components", [])
            print(f"S{sec}: 产品类型 + {len(comps)} 个成分")
        else:
            print(f"S{sec}: {len(items)} 个字段/子表项")
    print(f"保留模板结构：{payload.get('keep_structure', [])}")
    print(f"空值策略：{payload.get('empty_policy', 'preserve')}")


def _default_output(template: Path) -> Path:
    """生成带时间戳的默认输出路径。"""
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return template.with_name(f"{template.stem}_交互覆写_{stamp}.docx")


def run(args: argparse.Namespace) -> int:
    """执行交互或 JSON 驱动的完整流水线。"""
    template = Path(args.template).expanduser().resolve()
    if not template.is_file():
        raise FileNotFoundError(f"模板不存在：{template}")
    state = _load_state(template)
    sections = _parse_sections(args.sections)
    if args.input_json:
        payload = _load_input_json(Path(args.input_json), state)
    else:
        print("MSDS 17 节中文交互覆写表单")
        print(f"模板：{template}")
        print("说明：直接 Enter 保留模板值；输入 !清空 表示清空；多行字段以【结束】结束。")
        _edit_selected_sections(state, sections)
        payload = build_write_items(state)

    if not payload.get("sections"):
        print("没有需要覆写的字段，已取消。")
        return 2
    validate_write_items_permissions(payload)
    _preview(payload)

    write_items_path = Path(args.write_items).expanduser().resolve() if args.write_items else template.with_name(
        f"{template.stem}_write_items_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
    write_items_path.parent.mkdir(parents=True, exist_ok=True)
    write_items_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n写入项已保存：{write_items_path}")

    if not args.yes:
        answer = input("确认覆写并产出 Word？[Y/n] ").strip().lower()
        if answer not in {"", "y", "yes", "是", "确认"}:
            print("已取消覆写；写入项 JSON 保留。")
            return 0

    out_path = Path(args.out).expanduser().resolve() if args.out else _default_output(template)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    overwrite_engine.overwrite(
        str(template), payload, str(out_path),
        sections=set(sections),
        empty_policy=payload.get("empty_policy", "preserve"),
        missing_policy=payload.get("missing_policy", "preserve"),
        missing_text=payload.get("missing_text", ""),
        field_map=args.field_map,
    )
    ok, problems = overwrite_engine.verify_output(
        str(template), str(out_path), payload, sections=set(sections))
    if not ok:
        print("\n闭环校验失败：")
        for problem in problems:
            print(f"- {problem}")
        return 1
    print(f"\n覆写完成：{out_path}")
    print("闭环校验：通过")
    return 0


def build_parser() -> argparse.ArgumentParser:
    """构造 CLI 参数解析器。"""
    parser = argparse.ArgumentParser(description="MSDS 17 节中文交互表单与 Word 覆写")
    parser.add_argument("--template", default=str(DEFAULT_TEMPLATE), help="指定 MSDS 模板 docx")
    parser.add_argument("--out", help="产出 Word 路径；缺省在模板同目录生成时间戳文件")
    parser.add_argument("--write-items", help="写入项 JSON 路径；缺省自动生成")
    parser.add_argument("--input-json", help="跳过交互，直接使用已有 write_items JSON")
    parser.add_argument("--sections", default="all", help="交互节号：all、0-16 或 1,3,8,15")
    parser.add_argument("--field-map", default=None, help="字段映射 JSON")
    parser.add_argument("--yes", action="store_true", help="跳过确认，仅用于已审阅 JSON 的自动产出")
    return parser


def main() -> int:
    """CLI 主入口。"""
    try:
        return run(build_parser().parse_args())
    except (KeyboardInterrupt, EOFError):
        print("\n已取消。")
        return 130
    except Exception as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

__all__ = ["build_write_items", "main", "run"]
