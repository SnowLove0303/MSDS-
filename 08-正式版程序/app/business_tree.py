# -*- coding: utf-8 -*-
"""业务树下钻：总型号库（大类→型号→节→字段→值）、CAS 库（物质→别名→关联型号）。"""
from __future__ import annotations

import json
import sqlite3
from typing import Any

from . import lib_discovery


# ---------------------------------------------------------------------------
# 总型号库业务树
# ---------------------------------------------------------------------------
def categories(conn: sqlite3.Connection) -> list[dict[str, Any]]:
    """大类列表（category_parent 去重 + 型号数）。"""
    rows = conn.execute(
        "SELECT category_parent AS parent, COUNT(*) AS cnt "
        "FROM model_category GROUP BY category_parent ORDER BY parent").fetchall()
    return [{"parent": r["parent"], "model_count": r["cnt"]} for r in rows]


def models_by_category(conn: sqlite3.Connection, parent: str) -> list[dict[str, Any]]:
    """某大类下的型号列表。"""
    rows = conn.execute(
        "SELECT c.model_id, m.model, m.source, m.fields_count "
        "FROM model_category c JOIN msds_model m ON m.model_id = c.model_id "
        "WHERE c.category_parent = ? ORDER BY m.model", (parent,)).fetchall()
    return [{"model_id": r["model_id"], "model": r["model"],
             "source": r["source"], "fields_count": r["fields_count"]} for r in rows]


def model_sections(conn: sqlite3.Connection, model_id: int) -> list[dict[str, Any]]:
    """型号的 17 节概览（节号 + 字段数）。"""
    rows = conn.execute(
        "SELECT section, COUNT(*) AS cnt FROM msds_field "
        "WHERE model_id = ? GROUP BY section ORDER BY section", (model_id,)).fetchall()
    return [{"section": r["section"], "field_count": r["cnt"]} for r in rows]


def section_fields(conn: sqlite3.Connection, model_id: int, section: int) -> list[dict[str, Any]]:
    """某节的字段明细（父子级：seq/label/value/kind/sub_rows）。"""
    rows = conn.execute(
        "SELECT seq, label, value, std_name, kind, editable, row_index, sub_header, sub_rows "
        "FROM msds_field WHERE model_id = ? AND section = ? "
        "ORDER BY row_index, seq", (model_id, section)).fetchall()
    out = []
    for r in rows:
        item = {
            "seq": r["seq"], "label": r["label"], "value": r["value"],
            "std_name": r["std_name"], "kind": r["kind"], "editable": r["editable"],
        }
        if r["sub_header"]:
            try:
                item["sub_header"] = json.loads(r["sub_header"])
            except (json.JSONDecodeError, TypeError):
                item["sub_header"] = r["sub_header"]
        if r["sub_rows"]:
            try:
                item["sub_rows"] = json.loads(r["sub_rows"])
            except (json.JSONDecodeError, TypeError):
                item["sub_rows"] = r["sub_rows"]
        out.append(item)
    return out


def model_summary(conn: sqlite3.Connection, model_id: int) -> dict[str, Any] | None:
    row = conn.execute(
        "SELECT model_id, model, source, source_file, fields_count, created_at "
        "FROM msds_model WHERE model_id = ?", (model_id,)).fetchone()
    return dict(row) if row else None


# ---------------------------------------------------------------------------
# CAS 库业务树
# ---------------------------------------------------------------------------
def cas_substances(conn: sqlite3.Connection, keyword: str = "") -> list[dict[str, Any]]:
    sql = ("SELECT cas_id, cas_no, standard_name, category, registry_status, "
           "source_count, model_count FROM cas_substance")
    args: tuple = ()
    if keyword.strip():
        like = f"%{keyword.strip()}%"
        sql += " WHERE cas_no LIKE ? OR standard_name LIKE ?"
        args = (like, like)
    sql += " ORDER BY standard_name"
    rows = conn.execute(sql, args).fetchall()
    return [dict(r) for r in rows]


def cas_substance_detail(conn: sqlite3.Connection, cas_id: int,
                         msds_conn: sqlite3.Connection | None = None) -> dict[str, Any] | None:
    row = conn.execute(
        "SELECT cas_id, cas_no, standard_name, category, registry_status, "
        "source_count, model_count, remarks FROM cas_substance WHERE cas_id = ?",
        (cas_id,)).fetchone()
    if not row:
        return None
    detail = dict(row)
    aliases = conn.execute(
        "SELECT alias, alias_type, occurrence_count FROM cas_alias "
        "WHERE cas_id = ? ORDER BY occurrence_count DESC", (cas_id,)).fetchall()
    detail["aliases"] = [dict(a) for a in aliases]
    # cas_model_usage 存的是型号名字符串 (model)，msds_model 在总型号库中，
    # 因此跨库反查：先取关联型号名，再在总型号库按名匹配。
    usage_rows = conn.execute(
        "SELECT model, raw_name, raw_cas, concentration FROM cas_model_usage WHERE cas_id = ? "
        "ORDER BY model", (cas_id,)).fetchall()
    detail["models"] = [dict(u) for u in usage_rows]
    if msds_conn is not None and usage_rows:
        names = [u["model"] for u in usage_rows]
        placeholders = ",".join("?" for _ in names)
        resolved = msds_conn.execute(
            f"SELECT model_id, model, source FROM msds_model "
            f"WHERE model IN ({placeholders})", names).fetchall()
        resolved_map = {r["model"]: dict(r) for r in resolved}
        for u in detail["models"]:
            hit = resolved_map.get(u["model"])
            if hit:
                u["model_id"] = hit["model_id"]
    notes = conn.execute(
        "SELECT note_type, note, source FROM cas_note WHERE cas_no = ?",
        (row["cas_no"],)).fetchall()
    detail["notes"] = [dict(n) for n in notes]
    return detail


# ---------------------------------------------------------------------------
# 通用表查询（库→表→行→列）
# ---------------------------------------------------------------------------
def table_columns(conn: sqlite3.Connection, table: str) -> list[str]:
    rows = conn.execute(f'PRAGMA table_info("{table}")').fetchall()
    return [r["name"] for r in rows]


def table_rows(conn: sqlite3.Connection, table: str, page: int = 1, page_size: int = 50,
               keyword: str = "") -> dict[str, Any]:
    cols = table_columns(conn, table)
    where, args = "", []
    if keyword.strip():
        like = f"%{keyword.strip()}%"
        conds = []
        for c in cols[:8]:  # 只在前 8 列做关键词过滤，避免文本列拖慢
            conds.append(f'CAST("{c}" AS TEXT) LIKE ?')
        where = " WHERE " + " OR ".join(conds)
        args = [like] * len(conds)
    total = conn.execute(f'SELECT COUNT(*) FROM "{table}"{where}', args).fetchone()[0]
    off = (page - 1) * page_size
    rows = conn.execute(
        f'SELECT * FROM "{table}"{where} ORDER BY rowid LIMIT ? OFFSET ?',
        args + [page_size, off]).fetchall()
    return {
        "columns": cols,
        "total": total,
        "page": page,
        "page_size": page_size,
        "rows": [dict(r) for r in rows],
    }


def row_detail(conn: sqlite3.Connection, table: str, rowid: int) -> dict[str, Any] | None:
    row = conn.execute(
        f'SELECT * FROM "{table}" WHERE rowid = ?', (rowid,)).fetchone()
    return dict(row) if row else None


def resolve_library(key: str) -> dict[str, Any] | None:
    for lib in lib_discovery.discover_libraries():
        if lib["key"] == key:
            return lib
    return None
