# -*- coding: utf-8 -*-
"""MSDS 导入检查报告: 缺失节 / 必填字段 / 解析异常 / 重复检测 / 待归并标签.

在"导入 → 检查 → 添加 → 写入"流程中, 写入前先生成检查报告,
供 GUI 直观核对; 检查项均不阻断入库 (宁缺毋滥原则由用户决定),
但逐项列出, 便于发现源文件问题后编辑修正.
"""
from __future__ import annotations

from .catalog import extract_model_name, get_product, resolve_field_key


# 必填字段 (标准字段 key): {节: [field_key...]}
REQUIRED_FIELDS = {
    1: ["产品名称", "中文名称", "供应商名称", "供应商地址", "电话", "传真"],
    3: [],  # 成分用 components 计数单独判断
    4: ["一般措施", "误服", "接触眼睛", "接触皮肤", "吸入"],
    9: ["外观", "密度", "水溶性"],
    16: [],
}

# 必填字段别名 (兼容英文模板差异)
_REQ_ALIASES = {
    "产品名称": ["产品名称", "产品使用建议", "Trade name"],
    "供应商名称": ["供应商名称", "Name of supplier", "Name"],
}


def _has_field(field_keys: set[str], wanted: str) -> bool:
    """是否已有该标准字段 (含别名)."""
    if wanted in field_keys:
        return True
    for a in _REQ_ALIASES.get(wanted, []):
        if a in field_keys:
            return True
    return False


def check_doc(result, conn=None, *, category: str = "", model_name: str | None = None):
    """对一个 ParseResult 生成导入检查报告.

    conn 提供时额外检查: 重复检测 (sha256/型号) + 未归并标签(待确认池).
    """
    report = {
        "model": model_name or extract_model_name(result.file_name),
        "category": category,
        "missing_sections": sorted(n for n in range(1, 17) if n not in result.sections),
        "anomalies": [
            {"level": a.level, "section": a.section, "message": a.message}
            for a in result.anomalies],
        "fields_count": sum(len(s.fields) for s in result.sections.values()),
        "lines_count": sum(len(s.lines) for s in result.sections.values()),
        "components_count": len(result.sections.get(3).components) if result.sections.get(3) else 0,
        "pictograms_count": len(result.images),
        "duplicate_sha": None,
        "duplicate_model": None,
        "unmatched_labels": [],   # 字典未命中 → 待确认池 (独立列)
    }

    # 必填字段检查 (按节遍历字段, 用 resolve_field_key 归一到标准 key)
    present: dict[int, set[str]] = {}
    for num, sec in result.sections.items():
        s = set()
        for row in sec.iter_rows():
            if row.kind == "field" and row.label:
                k = resolve_field_key(conn, num, row.label) if conn else row.label
                s.add(k)
        present[num] = s
    missing_req = []
    for sec, wants in REQUIRED_FIELDS.items():
        for w in wants:
            if not _has_field(present.get(sec, set()), w):
                missing_req.append(f"S{sec}·{w}")
    if report["components_count"] == 0:
        missing_req.append("S3·成分")
    report["required_missing"] = missing_req

    # 待归并标签: 字典未命中且静态归一未命中 → 原样独立列
    if conn is not None:
        un = []
        seen = set()
        for num, sec in result.sections.items():
            for row in sec.iter_rows():
                if row.kind != "field" or not row.label:
                    continue
                k = resolve_field_key(conn, num, row.label)
                if k == row.label:  # 兜底原样 = 未归并
                    key = (num, k)
                    if key not in seen:
                        seen.add(key)
                        un.append({"section": num, "label": k})
        report["unmatched_labels"] = un

    # 重复检测
    if conn is not None:
        if result.sha256:
            dup = conn.execute("SELECT id, model_name FROM products "
                               "WHERE sha256=? AND active=1", (result.sha256,)).fetchone()
            if dup:
                report["duplicate_sha"] = {"id": dup["id"], "model": dup["model_name"]}
        if not report["duplicate_sha"]:
            cat_id = None
            if category:
                c = conn.execute("SELECT id FROM categories WHERE name=?",
                                 (category,)).fetchone()
                cat_id = c["id"] if c else None
            dup = conn.execute("SELECT id, model_name FROM products "
                               "WHERE category_id IS ? AND model_name=? AND active=1",
                               (cat_id, report["model"])).fetchone()
            if dup:
                report["duplicate_model"] = {"id": dup["id"], "model": dup["model_name"]}

    return report


def check_file(result, conn, category="", model_name=None) -> dict:
    """检查报告的便捷封装 (同 check_doc)."""
    return check_doc(result, conn, category=category, model_name=model_name)
