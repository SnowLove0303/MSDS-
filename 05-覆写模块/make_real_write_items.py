#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
真实 MSDS → 标准化写入项 JSON（批量流水线雏形）
============================================
用结构读取 read_msds 读取真实 MSDS 的三级结构（节→标题→字段），
转成覆写引擎可用的写入项 JSON。值原样保留，不做内容推断。

用法:
  python make_real_write_items.py <真实MSDS.docx> --out <写入项.json>
  python make_real_write_items.py <目录> --out <写入项目录>   # 批量
"""
import argparse
import json
import re
import sys
from pathlib import Path

STRUCTURE_READ_DIR = Path(r"F:\正式项目与模块化内容\冠志\MSDS\02-检索系统")
if str(STRUCTURE_READ_DIR) not in sys.path:
    sys.path.insert(0, str(STRUCTURE_READ_DIR))
try:
    from core.extract import read_msds
except Exception as e:
    read_msds = None
    print(f"[WARN] 无法导入结构读取: {e}")

FIELD_ALIASES = {"嗅觉阀值": "嗅觉阈值"}

def seq_key(seq):
    if not seq:
        return None
    parts = str(seq).split(".")
    nums = []
    for p in parts:
        if p.isdigit():
            nums.append(int(p))
        else:
            break
    return tuple(nums)

def norm_label(s):
    if not s:
        return ""
    s = str(s).replace(" ", "").replace("　", "").replace("\xa0", "")
    return FIELD_ALIASES.get(s, s)

def norm_field(text):
    if not text:
        return "", ""
    s = str(text).replace("　", " ").replace("\xa0", " ").strip()
    m = re.match(r"^(\d+(?:\.\d+)*)\s*(.*)$", s)
    if m and m.group(2):
        seq, rest = m.group(1), m.group(2).strip()
    else:
        seq, rest = "", s
    rest = re.sub(r"[：:，,。；;]+$", "", rest).strip()
    return seq, norm_label(rest)

def real_to_write_items(result):
    """ParseResult → 写入项 JSON dict。
    用 sd.iter_rows() 展开结构化行（含单格文本节的 lines 拆分），
    保证 S2 等单格多行节的字段也能进入写入项。
    """
    sec_objs = {}
    for sec_no, sd in result.sections.items():
        if sec_no == 0:
            # 页眉页脚：只保留可变字段，排除固定标题/自动页码
            # （物料安全数据表=固定标题，页码=域自动生成，均不可覆写）
            hf_items = []
            for sr in sd.iter_rows():
                if sr.kind not in ("field",):
                    continue
                if not sr.label:
                    continue
                nl = norm_label(sr.label)
                if nl in ("物料安全数据表", "页码"):
                    continue
                hf_items.append({"seq": "", "label": sr.label, "value": sr.value})
            sec_objs["0"] = hf_items
            continue
        if sd.is_component_table:
            sec_objs[str(sec_no)] = {
                "产品类型": "",
                "components": [],
            }
            for f in sd.fields:
                _s, label = norm_field(f.label)
                if label == "产品类型":
                    sec_objs[str(sec_no)]["产品类型"] = f.value
            for c in sd.components:
                sec_objs[str(sec_no)]["components"].append({
                    "name": c.name, "cas": c.cas, "conc": c.conc})
        else:
            items = []
            # 判断该节是“表格结构”还是“单格文本拆出”：
            #  - 表格结构（sd.fields 非空，如 S1/S9）：空值字段保留（模板有对应行则覆写为空）
            #  - 单格文本拆出（sd.fields 空 + sd.lines 非空，如 S2）：只保留有值字段
            #    （空值的 sub/field 是源文件结构骨架，如 '2.1 物质或混合物的分类' 标题，
            #      不参与内容覆写，避免覆盖模板同名字段的真实值）
            structured = bool(sd.fields)
            seen = set()
            last_field = None       # 单格文本节里最近收集的字段，用于吸收通栏续行
            for sr in sd.iter_rows():
                if sr.kind == "section":
                    continue                       # 节标题
                if sr.kind == "sub":
                    last_field = None              # 进入新块，通栏续行归到新块
                if sr.kind == "note":
                    # 单格文本节里的通栏行：并入当前块最后一个字段
                    # （如 S2 标签要素的续行 "1-丁氧基乙醇"），表格结构节不受影响。
                    if not structured and last_field is not None:
                        extra = str(sr.value or "").strip()
                        if extra:
                            last_field["value"] = (
                                last_field.get("value", "") + "\n" + extra).strip()
                    continue
                if not sr.label:
                    continue
                if not structured and not sr.value.strip():
                    continue                       # 单格文本节的空值骨架跳过
                key = (sr.seq, norm_label(sr.label))
                if key in seen:
                    continue                       # 去重（源 S2 防范说明重复等）
                seen.add(key)
                it = {"seq": sr.seq, "label": sr.label, "value": sr.value}
                items.append(it)
                if not structured:
                    last_field = it
            sec_objs[str(sec_no)] = items
    _apply_naming_rules(sec_objs)
    return {"sections": sec_objs}


def _apply_naming_rules(sec_objs):
    """中文名称补型号规则（以模板/表单惯例为准）：
    模板惯例：中文名称 = 中文 + 空格 + 型号（如 '水性羟基聚酯-丙烯酸分散体 PEA-4139'）。
    若源文件把 中文名称 与 产品名称(型号) 分开，合并到 中文名称。
    """
    s1 = sec_objs.get("1")
    if not isinstance(s1, list):
        return
    prod = next((it["value"] for it in s1 if it["label"] == "产品名称"), "").strip()
    for it in s1:
        if it["label"] == "中文名称":
            cur = it["value"].strip()
            # 型号形态：如 OS-1330 / BEK-750 / PEA-4139（短、含字母数字连字符）
            model = prod if re.match(r"^[A-Za-z]{1,6}-?\d{2,}$", prod) else ""
            if not model:
                continue
            # 中文名称未含型号 → 补上
            if model not in cur:
                it["value"] = f"{cur} {model}" if cur else model
        # 中文名称补型号后，产品名称仍保留原值（通道已建好）

def main():
    ap = argparse.ArgumentParser(description="真实 MSDS → 写入项 JSON")
    ap.add_argument("path", help="真实 MSDS docx 或目录")
    ap.add_argument("--out", required=True, help="输出 json 文件 或 目录")
    args = ap.parse_args()

    if read_msds is None:
        sys.exit("[FAIL] 结构读取工具不可用")
    src = Path(args.path)
    out = Path(args.out)

    if src.is_dir():
        out.mkdir(parents=True, exist_ok=True)
        for f in sorted(src.glob("*.docx")):
            r = read_msds(f)
            items = real_to_write_items(r)
            target = out / (f.stem + ".json")
            target.write_text(json.dumps(items, ensure_ascii=False, indent=1), encoding="utf-8")
            print(f"生成 {target.name}  (sections={len(items['sections'])})")
    else:
        r = read_msds(str(src))
        items = real_to_write_items(r)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps(items, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"生成 {out}  (sections={len(items['sections'])})")

if __name__ == "__main__":
    main()
