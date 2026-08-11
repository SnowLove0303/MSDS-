# -*- coding: utf-8 -*-
"""机器树中间格式: ParseResult -> document_tree (唯一 node_id + 完整上下文).

ChatGPT 批量入库 12 项标准 #5/#8/#11:
  - 每个节点有唯一 node_id (S3.ING.001 式), 保证可寻址
  - 可重复实体 (成分) 拆为独立 entity, 不再挤在 Value
  - 原子节点携带完整上下文路径 (parent_id 链: 节 -> 分组 -> 实体)
  - 字段双轨: value (归一化) + raw_value (原文), 供审计

node_id 规则:
  S{n}                     节 (n=0..16)
  S{n}.{seq}               子标题 sub (如 S8.1, 取原文序号)
  S{n}.GROUP.{m}           分组标题 (空值 field 引导其下子项)
  S{n}.{seq}.{m} / S{n}.F{m}   字段 (有/无序号)
  S{n}.NOTE.{m}            说明 note
  S{n}.ING.{i}             成分实体, 字段 S{n}.ING.{i}.{name|cas|conc}
"""
from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

# 并列防护类别 (眼睛/身体/呼吸) 结束分组标题作用域 (与 export_single 同源)
_GROUP_END_KW = ("眼睛", "身体", "呼吸", "听力", "面部", "头部", "皮肤")


def _meta(result) -> dict:
    prod = ver = rev = ""
    sec0 = result.sections.get(0)
    if sec0:
        for row in sec0.iter_rows():
            lab = (row.label or "").lower()
            if not row.value:
                continue
            if "产品名称" in lab or "product" in lab:
                prod = prod or row.value
            elif "version" in lab:
                ver = ver or row.value
            elif "修订" in lab or "revision" in lab:
                rev = rev or row.value
    return {
        "product_name": prod, "version": ver, "revision_date": rev,
        "header": result.header, "footer": result.footer,
    }


def build_document_tree(result) -> dict:
    """ParseResult -> 机器树 dict (唯一 node_id / 上下文路径 / raw_value)."""
    tree = []
    for n in sorted(result.sections):
        sec = result.sections[n]
        node = {
            "node_id": f"S{n}",
            "type": "section",
            "title": sec.full_title or sec.title,
            "section_no": n,
            "children": [],
        }
        if n == 3 and sec.component_header:
            node["component_header"] = sec.component_header
        # 父级链栈 (同 export_single Sheet2): sub/分组压栈, 字段挂栈顶
        parent_stack: list[dict] = [node]
        sec_rows = [x for x in sec.iter_rows() if x.kind != "section"]
        f_count = g_count = note_count = 0
        sub_seen: dict[str, int] = {}
        sub_no_seq = 0
        last_field = None
        for i, row in enumerate(sec_rows):
            nxt = sec_rows[i + 1] if i + 1 < len(sec_rows) else None
            is_group = (n != 0 and row.kind == "field"
                        and not (row.value or "").strip()
                        and nxt is not None and nxt.kind in ("field", "sub"))
            if row.kind == "sub":
                # 子标题: 同级并列 -> 回到节级再入栈
                if len(parent_stack) > 1:
                    parent_stack = [node]
                if row.seq:
                    # 防重复: 同 seq 第二次出现追加出现序号
                    sub_seen[row.seq] = sub_seen.get(row.seq, 0) + 1
                    cnt = sub_seen[row.seq]
                    nid = (f"S{n}.{row.seq}" if cnt == 1
                           else f"S{n}.{row.seq}.{cnt}")
                else:
                    sub_no_seq += 1
                    nid = f"S{n}.SUB.{sub_no_seq:03d}"
                sn = {"node_id": nid, "type": "subsection", "title": row.label,
                      "seq": row.seq, "children": []}
                parent_stack.append(sn)
                parent_stack[-2]["children"].append(sn)
            elif is_group:
                g_count += 1
                gn = {"node_id": f"S{n}.GROUP.{g_count:03d}",
                      "type": "group", "title": row.label, "children": []}
                parent_stack.append(gn)
                parent_stack[-2]["children"].append(gn)
            elif row.kind == "field":
                f_count += 1  # 所有字段 (含孤儿) 占用唯一计数
                if not (row.label or "").strip():
                    # 孤儿节点 (无标签) -> 挂最近字段/当前父级
                    parent = last_field or parent_stack[-1]
                    if "children" not in parent:
                        parent["children"] = []  # 字段节点可承载孤儿内容
                    parent["children"].append({
                        "node_id": _field_id(n, row.seq, f_count),
                        "type": "field", "label": "", "seq": row.seq,
                        "value": row.value, "raw_value": "",
                        "editable": row.editable, "parent_id": parent["node_id"],
                    })
                    continue
                # 并列防护类别在分组标题后 -> 回到分组父级
                if (row.label.startswith(_GROUP_END_KW)
                        and len(parent_stack) > 2):
                    parent_stack = parent_stack[:-1]
                parent = parent_stack[-1]
                fn = {"node_id": _field_id(n, row.seq, f_count),
                      "type": "field", "label": row.label, "seq": row.seq,
                      "value": row.value, "raw_value": "",
                      "editable": row.editable, "parent_id": parent["node_id"]}
                parent["children"].append(fn)
                last_field = fn
            elif row.kind == "note":
                note_count += 1
                parent_stack[-1]["children"].append({
                    "node_id": f"S{n}.NOTE.{note_count:03d}",
                    "type": "text", "value": row.value,
                    "editable": row.editable, "parent_id": parent_stack[-1]["node_id"],
                })
        # 成分实体化
        if n == 3 and sec.components:
            node["entities"] = []
            for i, c in enumerate(sec.components, 1):
                ing = {
                    "node_id": f"S3.ING.{i:03d}",
                    "type": "entity",
                    "entity_type": "ingredient",
                    "fields": [
                        {"node_id": f"S3.ING.{i:03d}.name", "label": "名称",
                         "value": c.name, "raw_value": c.raw_name,
                         "parent_id": f"S3.ING.{i:03d}"},
                        {"node_id": f"S3.ING.{i:03d}.cas", "label": "CAS",
                         "value": c.cas, "raw_value": c.raw_cas,
                         "parent_id": f"S3.ING.{i:03d}"},
                        {"node_id": f"S3.ING.{i:03d}.conc", "label": "含量",
                         "value": c.conc, "raw_value": c.raw_conc,
                         "parent_id": f"S3.ING.{i:03d}"},
                    ],
                }
                node["entities"].append(ing)
        tree.append(node)

    return {
        "schema": "msds-document-tree/v1",
        "document": {
            "file_name": result.file_name,
            "file_path": result.file_path,
            "sha256": result.sha256,
            "sections_count": len(result.sections),
        },
        "meta": _meta(result),
        "tree": tree,
    }


def _field_id(n, seq, count) -> str:
    if seq:
        return f"S{n}.{seq}.{count}"
    return f"S{n}.F{count:03d}"


def export_document_tree(result, out_path) -> dict:
    """导出 document_tree.json, 返回节点统计."""
    import json
    tree = build_document_tree(result)
    out_path = Path(out_path)
    out_path.write_text(json.dumps(tree, ensure_ascii=False, indent=2),
                        encoding="utf-8")
    n_nodes = n_fields = n_entities = 0
    for sec in tree["tree"]:
        n_nodes += 1
        n_entities += len(sec.get("entities", []))

        def walk(children):
            nonlocal n_nodes, n_fields
            for c in children:
                n_nodes += 1
                if c["type"] == "field":
                    n_fields += 1
                walk(c.get("children", []))
        walk(sec.get("children", []))
    return {"path": str(out_path), "sections": len(tree["tree"]),
            "nodes": n_nodes, "fields": n_fields, "entities": n_entities}
