#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MSDS 数据库检索与覆写流水线
============================================
从 SQLite 数据库 (msds_standard.db) 检索指定型号数据，
自动构建符合覆写引擎契约的标准写入项 (write_items)，
调用 msds_overwrite_engine 注入标准 Word 模板并输出新的 MSDS docx 文档。

用法:
  python run_db_overwrite.py --model EC-1801
  python run_db_overwrite.py --model PEA-4139 --out outputs/PEA-4139_out.docx
"""

import os
import sys
import json
import argparse
from pathlib import Path

# 添加结构读取与覆写引擎到 sys.path
BASE_DIR = Path(os.environ.get(
    "MSDS_ROOT", str(Path(__file__).resolve().parents[1])))
READ_DIR = BASE_DIR / "02-检索系统"
ENGINE_DIR = BASE_DIR / "05-覆写模块"

if str(READ_DIR) not in sys.path:
    sys.path.insert(0, str(READ_DIR))
if str(ENGINE_DIR) not in sys.path:
    sys.path.insert(0, str(ENGINE_DIR))

import core.msds_db as m
import msds_overwrite_engine as moe

DB_PATH = Path(os.environ.get(
    "MSDS_DB_PATH",
    str(BASE_DIR / "03-数据库" / "正式库" / "Data Base" / "msds_standard.db")))
DEFAULT_TEMPLATE = BASE_DIR / "03-数据库" / "正式库" / "推导方案" / (
    "PEA-4139 MSDS_CN 冠志 模板.docx")
DEFAULT_FIELD_MAP = ENGINE_DIR / "field_maps_pea4139_cn.json"


def _require_canonical_template(path: Path) -> Path:
    """Enforce the single approved Chinese overwrite template."""
    canonical = DEFAULT_TEMPLATE.resolve()
    actual = path.resolve()
    if actual != canonical:
        raise ValueError(
            f"中文覆写只允许使用唯一模板 {canonical}，收到: {actual}")
    if not actual.is_file():
        raise FileNotFoundError(f"唯一中文覆写模板不存在: {actual}")
    return actual


def build_write_items_from_db(conn, model_id: int) -> dict:
    """从数据库中提取指定 model_id 的全节数据并构建 write_items。"""
    detail = m.model_detail(conn, model_id)
    model_name = detail.get("model", "")

    sections = {}

    # S0: 页眉与页脚
    s0_rows = m.listed_section_rows(conn, model_id, 0)
    s0_items = []
    for r in s0_rows:
        if r.kind == "field" and r.value and r.label not in ("物料安全数据表", "页码"):
            s0_items.append({"seq": "", "label": r.label, "value": r.value})
    if not s0_items:
        s0_items = [
            {"seq": "", "label": "Version", "value": "1.0"},
            {"seq": "", "label": "产品名称", "value": model_name},
            {"seq": "", "label": "公司名称", "value": "广州冠志新材料科技有限公司"},
            {"seq": "", "label": "产品型号", "value": model_name},
            {"seq": "", "label": "修订日期", "value": "2026-08-14"},
        ]
    sections["0"] = s0_items

    # S1 ~ S16
    for sec in range(1, 17):
        rows = m.listed_section_rows(conn, model_id, sec)
        if sec == 3:
            s3_type = "混合物"
            comps = []
            for r in rows:
                if r.kind == "field" and r.label == "产品类型" and r.value:
                    s3_type = r.value
                elif r.kind == "subtable" and r.label == "成分":
                    for row in r.sub_rows:
                        if len(row) >= 3 and any(str(x).strip() for x in row):
                            name_clean = str(row[0]).strip().replace(" ", "")
                            comps.append({
                                "name": name_clean,
                                "cas": str(row[1]).strip(),
                                "conc": str(row[2]).strip(),
                            })
            sections["3"] = {"产品类型": s3_type, "components": comps}
        else:
            items = []
            for r in rows:
                if r.kind == "field" and m.is_meaningful_value(r.value):
                    items.append({"seq": r.seq or "", "label": r.label, "value": r.value})
                elif r.kind == "note" and m.is_meaningful_value(r.value):
                    note_label = r.label
                    # S8 legacy note rows carry values without labels; map only
                    # unambiguous semantic notes to their fixed template fields.
                    if sec == 8 and not note_label:
                        text = str(r.value)
                        if text.strip().startswith("穿着适当的防护服"):
                            note_label = "身体防护"
                        elif "戴眼罩" in text or "戴护目镜" in text:
                            note_label = "眼睛防护"
                        elif text.strip().startswith("污染的手套"):
                            note_label = "建议"
                    if note_label:
                        items.append({"seq": r.seq or "", "label": note_label, "value": r.value})
            if sec == 8:
                # listed_section_rows intentionally hides untyped legacy notes;
                # recover only the three unambiguous S8 semantic notes here.
                raw_notes = conn.execute(
                    "SELECT value FROM msds_field WHERE model_id=? AND section=8 "
                    "AND kind='note' AND value IS NOT NULL AND value<>'' ORDER BY row_index,id",
                    (model_id,)).fetchall()
                existing_labels = {moe.norm_label(x.get("label", "")) for x in items}
                for (note_value,) in raw_notes:
                    text = str(note_value)
                    note_label = None
                    if text.strip().startswith("穿着适当的防护服"):
                        note_label = "身体防护"
                    elif "戴眼罩" in text or "戴护目镜" in text:
                        note_label = "眼睛防护"
                    elif text.strip().startswith("污染的手套"):
                        note_label = "建议"
                    if note_label and moe.norm_label(note_label) not in existing_labels:
                        items.append({"seq": "", "label": note_label, "value": text})
                        existing_labels.add(moe.norm_label(note_label))
            if items or sec in (1, 2, 9):
                sections[str(sec)] = items

    write_items = {
        "sections": sections,
        # 模板骨架默认只读；S9 由专用瘦身规则处理，其余节保留模板行。
        "keep_structure": [1, 2, 3, 4, 5, 6, 7, 8, 10, 11, 12, 13, 14, 15, 16],
        "empty_policy": "warn",
    }
    # 唯一输入口：归一化句内视觉换行、按标准字段合并并执行字段隔离。
    write_items = moe.prepare_write_items(write_items)
    moe.audit_input_content_isolation(write_items)
    return write_items


def run_pipeline(model_query: str = "EC-1801", template_path: str = None, out_path: str = None):
    if not DB_PATH.exists():
        raise FileNotFoundError(f"数据库文件不存在: {DB_PATH}")

    template_file = _require_canonical_template(
        Path(template_path) if template_path else DEFAULT_TEMPLATE)
    print(f"[模板] 唯一中文模板: {template_file}")

    conn = m.open_db(str(DB_PATH))
    print(f"[1/4] 连接数据库: {DB_PATH}")
    
    # 检索型号
    matches = m.find_models(conn, model_query)
    if not matches:
        all_models = m.list_models(conn)
        avail = [row[0] for row in all_models]
        raise ValueError(f"未在数据库中找到型号 '{model_query}'。当前库内可用型号: {avail}")
    
    target_model = matches[0]
    model_id = target_model[0]
    model_name = target_model[1]
    print(f"[2/4] 检索到目标型号: ID={model_id}, 型号={model_name}, 来源={target_model[2]}, 字段数={target_model[4]}")

    # 构建写入项
    print(f"[3/4] 正在从数据库提取并转换 16 节写入项...")
    write_items = build_write_items_from_db(conn, model_id)
    sec_count = len(write_items["sections"])
    print(f"      成功生成 {sec_count} 个 Section 写入项")

    # 确定输出路径
    if out_path:
        out_file = Path(out_path)
    else:
        out_dir = ENGINE_DIR / "outputs"
        out_dir.mkdir(parents=True, exist_ok=True)
        out_file = out_dir / f"{model_name}_MSDS_CN_冠志_覆写输出.docx"

    out_file.parent.mkdir(parents=True, exist_ok=True)

    # 执行覆写
    print(f"[4/4] 正在调用覆写引擎注入模板: {template_file.name} -> {out_file.name} ...")
    field_map_arg = str(DEFAULT_FIELD_MAP) if DEFAULT_FIELD_MAP.exists() else None
    logs = moe.overwrite(
        str(template_file),
        write_items,
        str(out_file),
        field_map=field_map_arg,
        missing_policy="no_data",
        missing_text="无数据",
    )
    print(f"      覆写完成，记录操作日志 {len(logs)} 条")

    # 闭环校验
    ok, probs = moe.verify_output(str(template_file), str(out_file), write_items)
    if ok:
        print(f"      [PASS] 闭环校验通过！所有字段与成分表一致。")
    else:
        print(f"      [WARN] 闭环校验发现 {len(probs)} 处差异:")
        for p in probs[:10]:
            print(f"        - {p}")

    print(f"\n>> 最终生成文档: {out_file}")
    return str(out_file)


def main():
    parser = argparse.ArgumentParser(description="MSDS 数据库检索与覆写生成")
    parser.add_argument("--model", default="EC-1801", help="目标型号名称 (默认: EC-1801)")
    parser.add_argument("--template", default=None, help="底模 docx 路径")
    parser.add_argument("--out", default=None, help="输出 docx 路径")
    args = parser.parse_args()

    run_pipeline(model_query=args.model, template_path=args.template, out_path=args.out)


if __name__ == "__main__":
    main()
