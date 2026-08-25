# -*- coding: utf-8 -*-
"""MSDS 正式版 Web 服务：只编排既有数据库、表单和覆写模块。"""
from __future__ import annotations

import base64
import json
import mimetypes
import re
import shutil
import sys
import tempfile
import traceback
import zipfile
from datetime import datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, urlparse

ROOT = Path(r"F:\正式项目与模块化内容\冠志\MSDS")
WEB_ROOT = ROOT / "08-正式版程序"
DB_PATH = ROOT / "03-数据库" / "正式库" / "Data Base" / "msds_standard.db"
CAS_DB_PATH = ROOT / "03-数据库" / "正式库" / "Data Base" / "cas_library.db"
CAS_RESULT_DB_PATH = ROOT / "03-数据库" / "正式库" / "Data Base" / "cas_section2_result.db"
CAS_MAPPING_DB_PATH = ROOT / "03-数据库" / "正式库" / "Data Base" / "cas_field_mapping.db"
REACH_DB_PATH = ROOT / "04-推断引擎" / "正式法律法规库" / "物质限制清单" / "REACH-SVHC-253项-物质数据库-2026-02-04.db"
ROHS_DB_PATH = ROOT / "04-推断引擎" / "正式法律法规库" / "物质限制清单" / "EU-RoHS-2011-65-EU-2015-863-限制物质数据库.db"
ANNEX_XVII_DB_PATH = ROOT / "04-推断引擎" / "正式法律法规库" / "物质限制清单" / "REACH-Annex-XVII-限制物质数据库.db"
HSF_001_DB_PATH = ROOT / "04-推断引擎" / "正式法律法规库" / "物质限制清单" / "HSF-001-有害物质清单.db"
DRAFT_ROOT = ROOT / "_codex_work" / "web_drafts"
sys.path.insert(0, str(ROOT / "02-检索系统"))
sys.path.insert(0, str(ROOT / "05-覆写模块"))
sys.path.insert(0, str(WEB_ROOT / "CAS API 搜索"))

from core.form_schema import build_form_schema, form_to_write_items, ComponentRow  # noqa: E402
from core.msds_db import (find_models, listed_rows_from_result,
                          listed_section_rows, model_detail,
                          model_to_write_items, open_db, wide_row)  # noqa: E402
from core.docx_reader import read_msds  # noqa: E402
from core.extract import build_hierarchy  # noqa: E402
import msds_overwrite_engine as overwrite_engine  # noqa: E402
import msds_form_overwrite as form_overwrite  # noqa: E402
import msds_multiformat as multiformat  # noqa: E402
import tds_overwrite_web as tds_web  # noqa: E402
import cas_query as cas_online_query  # noqa: E402


CANONICAL_TEMPLATE = ROOT / "03-数据库" / "正式库" / "推导方案" / "PEA-4139 MSDS_CN 冠志 模板.docx"
CANONICAL_TEMPLATE_EN = ROOT / "03-数据库" / "正式库" / "推导方案" / "PEA-4139 模板 EN.docx"
WEB_OVERWRITE_ROOT = WEB_ROOT / "outputs"
GHS_PICTOGRAM_ASSET_DIR = ROOT / "04-推断引擎" / "判断skill" / "templates" / "pictograms"
GHS_PICTOGRAMS = {
    "GHS01": "爆炸物", "GHS02": "易燃", "GHS03": "氧化剂",
    "GHS04": "加压气体", "GHS05": "腐蚀", "GHS06": "急性毒性",
    "GHS07": "感叹号", "GHS08": "健康危害", "GHS09": "环境危害",
}


def _ghs_pictogram_catalog():
    """Expose the database-backed GHS registry and icon bytes to the same-origin UI."""
    import sqlite3

    items = {}
    try:
        uri = f"file:{CAS_RESULT_DB_PATH.as_posix()}?mode=ro"
        with sqlite3.connect(uri, uri=True) as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                "SELECT pictogram_code, label_zh, mime_type, image_blob "
                "FROM cas_s2_pictogram_catalog ORDER BY pictogram_code"
            ).fetchall()
        for row in rows:
            encoded = base64.b64encode(row["image_blob"] or b"").decode("ascii")
            items[row["pictogram_code"]] = {
                "code": row["pictogram_code"],
                "label": row["label_zh"],
                "data_uri": f"data:{row['mime_type']};base64,{encoded}" if encoded else "",
            }
        if items:
            return items
    except sqlite3.Error:
        pass
    for code, label in GHS_PICTOGRAMS.items():
        asset = GHS_PICTOGRAM_ASSET_DIR / f"{code}.png"
        item = {"code": code, "label": label, "data_uri": ""}
        if asset.is_file():
            encoded = base64.b64encode(asset.read_bytes()).decode("ascii")
            item["data_uri"] = f"data:image/png;base64,{encoded}"
        items[code] = item
    return items


def _jsonable(value):
    if hasattr(value, "__dict__"):
        return {k: _jsonable(v) for k, v in value.__dict__.items()}
    if isinstance(value, (list, tuple)):
        return [_jsonable(v) for v in value]
    if isinstance(value, dict):
        return {str(k): _jsonable(v) for k, v in value.items()}
    return value


def _node_to_dict(node):
    """Serialize the existing GUI hierarchy for the Web import view."""
    out = {}
    for attr in ("number", "title", "full_title", "seq", "label", "value",
                 "std_name", "kind", "editable", "index", "unit"):
        if hasattr(node, attr):
            value = getattr(node, attr)
            if value is not None and not callable(value):
                out[attr] = value
    for attr in ("big_titles", "children", "direct_fields"):
        if hasattr(node, attr):
            out[attr] = [_node_to_dict(item) for item in getattr(node, attr)]
    if hasattr(node, "rows") and isinstance(node.rows, list):
        out["rows"] = [_jsonable(row) for row in node.rows]
    return out


def _schema():
    s = build_form_schema()
    return {"s1": [f.to_dict() for f in s.s1],
            "s3": [f.to_dict() for f in s.s3],
            "s9": [f.to_dict() for f in s.s9]}


def _rows(conn, model_id: int, section: int | None = None):
    # The GUI contract is a fixed S0~S16 tree: S0 is the header/footer
    # section, and S1~S16 are the standard MSDS sections.
    sections = [section] if section is not None else list(range(17))
    result = {}
    for sec in sections:
        result[str(sec)] = []
        for row in listed_section_rows(conn, model_id, sec):
            result[str(sec)].append(_jsonable(row))
    return result


def _form_from_model(conn, model_id: int):
    detail = model_detail(conn, model_id)
    values = {"s1": {}, "s3": {"产品类型": "混合物", "components": []}, "s9": {}}
    for row in listed_section_rows(conn, model_id, 1):
        if row.label:
            values["s1"][row.label] = row.value
    for row in listed_section_rows(conn, model_id, 3):
        if row.kind == "field" and row.label == "产品类型":
            values["s3"]["产品类型"] = row.value
        elif row.kind == "subtable":
            values["s3"]["components"] = [
                {"name": str(x[0]) if len(x) > 0 else "",
                 "cas": str(x[1]) if len(x) > 1 else "",
                 "conc": str(x[2]) if len(x) > 2 else ""}
                for x in row.sub_rows]
    for row in listed_section_rows(conn, model_id, 9):
        if row.label:
            values["s9"][row.label] = row.value
    return {"detail": detail, "values": values}


def _library_info():
    items = []
    with open_db(DB_PATH) as conn:
        models, fields, categories = conn.execute(
            "SELECT (SELECT COUNT(*) FROM msds_model), (SELECT COUNT(*) FROM msds_field), "
            "(SELECT COUNT(*) FROM model_category)").fetchone()
        items.append({"id": "models", "name": "中文总型号库", "kind": "型号库",
                      "path": str(DB_PATH), "stats": {"型号": models, "字段": fields, "归类": categories},
                      "description": "型号唯一；Section 1 / 3 / 9 与标准骨架字段集中检索。"})
        mapping, unmapped, schema = conn.execute(
            "SELECT (SELECT COUNT(*) FROM field_mapping), (SELECT COUNT(*) FROM msds_unmapped), "
            "(SELECT COUNT(*) FROM schema_field)").fetchone()
        items.append({"id": "field_mapping", "name": "字段映射库", "kind": "规则库",
                      "path": str(DB_PATH), "stats": {"映射": mapping, "待归类": unmapped, "标准字段": schema},
                      "description": "原始标签、标准字段名、别名和 Section 9 映射规则。"})
    if CAS_DB_PATH.exists():
        import sqlite3
        with sqlite3.connect(CAS_DB_PATH) as conn:
            def count(table):
                try:
                    return conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
                except sqlite3.Error:
                    return 0
        items.append({"id": "cas", "name": "CAS 库", "kind": "物质库", "path": str(CAS_DB_PATH),
                          "stats": {"物质": count("cas_substance"),
                                    "确认 CAS": conn.execute("SELECT COUNT(*) FROM cas_substance WHERE TRIM(cas_no)<>''").fetchone()[0],
                                    "型号使用关系": count("cas_model_usage"), "别名": count("cas_alias")},
                          "description": "成分、CAS、别名和型号使用关系。"})
    if CAS_MAPPING_DB_PATH.exists():
        import sqlite3
        with sqlite3.connect(CAS_MAPPING_DB_PATH) as conn:
            def mapping_count(table):
                try:
                    return conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
                except sqlite3.Error:
                    return 0
            items.append({"id": "cas_mapping", "name": "CAS 字段映射库", "kind": "映射库",
                          "path": str(CAS_MAPPING_DB_PATH),
                          "stats": {"身份": mapping_count("cas_identity"),
                                    "字段映射": mapping_count("cas_field_mapping"),
                                    "观察记录": mapping_count("cas_mapping_observation"),
                                    "冲突": mapping_count("cas_mapping_conflict")},
                          "description": "CAS/无 CAS 成分规范名、原始名称、异写和冲突关系。"})
    items.append({"id": "categories", "name": "型号大类索引", "kind": "索引库",
                  "path": r"F:\冠志工作空间\产品\TDS MSDS\TDS MSDS\产品 TDS MSDS -- WORD版本",
                  "stats": {"父级": len(_category_tree()), "型号": sum(len(x["models"]) for x in _category_tree())},
                  "description": "按指定 WORD 版本目录父子级结构生成，父级为产品类别，子级为型号。"})
    return items


def _category_tree():
    with open_db(DB_PATH) as conn:
        rows = conn.execute("SELECT category_parent, category_child, model_id FROM model_category ORDER BY category_parent, category_child").fetchall()
        groups = {}
        for parent, child, model_id in rows:
            groups.setdefault(parent or "未归类", {"name": parent or "未归类", "models": []})["models"].append({"id": model_id, "model": child})
        return list(groups.values())


def _cas_search(query: str = "", status: str = ""):
    """CAS 库检索：复用正式 cas_library.db，返回物质及其型号使用关系。"""
    import sqlite3

    q = str(query or "").strip()
    with sqlite3.connect(CAS_DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        params = []
        clauses = []
        if q:
            like = f"%{q}%"
            clauses.append("(c.cas_no LIKE ? OR c.standard_name LIKE ? OR c.category LIKE ? " \
                           "OR c.registry_status LIKE ? OR c.remarks LIKE ? " \
                           "OR EXISTS (SELECT 1 FROM cas_alias a WHERE a.cas_id=c.cas_id AND a.alias LIKE ?) " \
                           "OR EXISTS (SELECT 1 FROM cas_model_usage u WHERE u.cas_id=c.cas_id " \
                           "AND (u.model LIKE ? OR u.raw_name LIKE ? OR u.raw_cas LIKE ?)))")
            params.extend([like] * 9)
        if status == "has_cas":
            clauses.append("NULLIF(TRIM(c.cas_no), '') IS NOT NULL")
        elif status == "no_cas":
            clauses.append("NULLIF(TRIM(c.cas_no), '') IS NULL")
        elif status == "secret":
            clauses.append("(c.registry_status LIKE '%商业机密%' OR c.category LIKE '%保密%' "
                           "OR c.standard_name LIKE '%商业机密%')")
        where = ("WHERE " + " AND ".join(clauses)) if clauses else ""
        substances = conn.execute(
            "SELECT c.cas_id, c.cas_no, c.standard_name, c.category, c.registry_status, "
            "c.source_count, c.model_count, c.remarks FROM cas_substance c "
            f"{where} ORDER BY CASE WHEN c.cas_no=? AND c.cas_no<>'' THEN 0 ELSE 1 END, c.standard_name",
            params + [q]).fetchall()
        items = []
        for c in substances:
            aliases = [r[0] for r in conn.execute(
                "SELECT alias FROM cas_alias WHERE cas_id=? ORDER BY alias_id", (c["cas_id"],)).fetchall()]
            usage = [dict(r) for r in conn.execute(
                "SELECT model, raw_name, raw_cas, concentration, source_file "
                "FROM cas_model_usage WHERE cas_id=? ORDER BY model, usage_id", (c["cas_id"],)).fetchall()]
            items.append({**dict(c), "aliases": aliases, "usage": usage,
                          "s2_result": _cas_s2_result(c["cas_no"])})
        return items


def _reach_svhc_check(identifier: str):
    """Read-only REACH SVHC CAS/EC check plus the no-CAS review queue."""
    import sqlite3

    value = str(identifier or "").strip()
    if not value:
        raise ValueError("请输入 CAS 号或 EC 号")
    if not re.fullmatch(r"\d{2,7}-\d{2,7}-\d", value):
        raise ValueError("CAS 号或 EC 号格式不正确")
    if not REACH_DB_PATH.is_file():
        raise FileNotFoundError(f"REACH 数据库不存在：{REACH_DB_PATH}")

    uri = f"file:{REACH_DB_PATH.as_posix()}?mode=ro"
    with sqlite3.connect(uri, uri=True) as conn:
        conn.row_factory = sqlite3.Row
        matched = []
        for table in ("reach_svhc_main", "reach_svhc_missing_identifier"):
            for row in conn.execute(
                f'SELECT "中文名称", "英文名称", "物质描述", "EC号", "CAS号" '
                f'FROM {table} ORDER BY rowid'
            ):
                cas_identifiers = re.findall(r"\d{2,7}-\d{2,7}-\d", str(row["CAS号"] or ""))
                ec_identifiers = re.findall(r"\d{2,7}-\d{2,7}-\d", str(row["EC号"] or ""))
                if value in cas_identifiers or value in ec_identifiers:
                    matched.append(dict(row))
        no_cas = [
            {"中文名称": row[0], "英文名称": row[1]}
            for row in conn.execute(
                'SELECT "中文名称", "英文名称" FROM reach_svhc_missing_identifier '
                'WHERE trim(coalesce("CAS号", "")) = "" ORDER BY rowid'
            )
        ]
    return {
        "query": value,
        "matched": bool(matched),
        "records": matched,
        "no_cas_pending": no_cas,
        "no_cas_count": len(no_cas),
    }


def _reach_svhc_list():
    """Return the complete 253-entry REACH list for the read-only table page."""
    import sqlite3

    if not REACH_DB_PATH.is_file():
        raise FileNotFoundError(f"REACH 数据库不存在：{REACH_DB_PATH}")
    uri = f"file:{REACH_DB_PATH.as_posix()}?mode=ro"
    rows = []
    with sqlite3.connect(uri, uri=True) as conn:
        conn.row_factory = sqlite3.Row
        for table in ("reach_svhc_main", "reach_svhc_missing_identifier"):
            for row in conn.execute(
                f'SELECT "中文名称", "英文名称", "CAS号", "EC号", "物质描述" '
                f'FROM {table} ORDER BY rowid'
            ):
                rows.append({
                    "中文名称": row["中文名称"] or "",
                    "英文名称": row["英文名称"] or "",
                    "CAS号": row["CAS号"] or "",
                    "EC号": row["EC号"] or "",
                    "物质描述": row["物质描述"] or "",
                })
    return {"count": len(rows), "columns": ["中文名称", "英文名称", "CAS号", "EC号", "物质描述"], "rows": rows}


def _rohs_check(identifier: str):
    """Read-only RoHS CAS/EC check from the formal material-level register."""
    import sqlite3

    value = str(identifier or "").strip()
    if not value:
        raise ValueError("请输入 CAS 号或 EC 号")
    if not re.fullmatch(r"\d{2,7}-\d{2,7}-\d", value):
        raise ValueError("CAS 号或 EC 号格式不正确")
    if not ROHS_DB_PATH.is_file():
        raise FileNotFoundError(f"RoHS 数据库不存在：{ROHS_DB_PATH}")

    uri = f"file:{ROHS_DB_PATH.as_posix()}?mode=ro"
    with sqlite3.connect(uri, uri=True) as conn:
        conn.row_factory = sqlite3.Row
        records = []
        for row in conn.execute(
            'SELECT "序号", "物质类别", "中文名称", "英文名称", "CAS号", "EC号", '
            '"均质材料限值", "限值ppm", "材料筛查说明", "法规依据" '
            'FROM rohs_restricted_substance ORDER BY "序号"'
        ):
            cas_identifiers = re.findall(r"\d{2,7}-\d{2,7}-\d", str(row["CAS号"] or ""))
            ec_identifiers = re.findall(r"\d{2,7}-\d{2,7}-\d", str(row["EC号"] or ""))
            if value in cas_identifiers or value in ec_identifiers:
                records.append(dict(row))
    return {"query": value, "matched": bool(records), "records": records}


def _hsf_001_check(identifier: str):
    """Read-only HSF 001 substance/CAS screening from the formal register."""
    import sqlite3

    value = str(identifier or "").strip()
    if not value:
        raise ValueError("请输入 CAS 号")
    if not re.fullmatch(r"\d{2,7}-\d{2,7}-\d", value):
        raise ValueError("CAS 号格式不正确")
    if not HSF_001_DB_PATH.is_file():
        raise FileNotFoundError(f"HSF 001 数据库不存在：{HSF_001_DB_PATH}")

    uri = f"file:{HSF_001_DB_PATH.as_posix()}?mode=ro"
    with sqlite3.connect(uri, uri=True) as conn:
        conn.row_factory = sqlite3.Row
        meta = {
            row["meta_key"]: row["meta_value"]
            for row in conn.execute("SELECT meta_key, meta_value FROM hsf_001_meta")
        }
        records = []
        rows = conn.execute(
            'SELECT "序号", "源表行号", "类别", "中文名称", "英文名称", "CAS号", "源表CAS号", '
            '"CAS映射状态", "物质级判定", "法规依据", "备注" '
            'FROM hsf_001_substance ORDER BY "序号"'
        ).fetchall()
        for row in rows:
            if value in re.findall(r"\d{2,7}-\d{2,7}-\d", str(row["CAS号"] or "")):
                records.append(dict(row))
        no_cas_count = sum(not str(row["CAS号"] or "").strip() for row in rows)
    return {
        "query": value,
        "matched": bool(records),
        "status": "不符合" if records else "符合（当前CAS未命中）",
        "records": records,
        "source": meta.get("来源", ""),
        "source_note": meta.get("来源说明", ""),
        "scope_note": meta.get("判定范围", ""),
        "no_cas_count": no_cas_count,
    }


def _hsf_001_list():
    """Return the 28 image rows and their current CAS mapping status."""
    import sqlite3

    if not HSF_001_DB_PATH.is_file():
        raise FileNotFoundError(f"HSF 001 数据库不存在：{HSF_001_DB_PATH}")
    uri = f"file:{HSF_001_DB_PATH.as_posix()}?mode=ro"
    columns = ["序号", "源表行号", "类别", "中文名称", "英文名称", "CAS号", "源表CAS号", "CAS映射状态", "物质级判定", "法规依据", "备注"]
    with sqlite3.connect(uri, uri=True) as conn:
        conn.row_factory = sqlite3.Row
        meta = {
            row["meta_key"]: row["meta_value"]
            for row in conn.execute("SELECT meta_key, meta_value FROM hsf_001_meta")
        }
        rows = [
            {column: row[column] or "" for column in columns}
            for row in conn.execute(
                'SELECT "序号", "源表行号", "类别", "中文名称", "英文名称", "CAS号", "源表CAS号", '
                '"CAS映射状态", "物质级判定", "法规依据", "备注" '
                'FROM hsf_001_substance ORDER BY "序号"'
            )
        ]
    return {"count": len(rows), "columns": columns, "rows": rows, "meta": meta}


def _compliance_check(identifier: str):
    """Run the same CAS/EC identifier through every currently connected register."""
    value = str(identifier or "").strip()
    hsf_result = (_hsf_001_check(value)
                  if re.fullmatch(r"\d{2,7}-\d{2,7}-\d", value)
                  else {
                      "query": value,
                      "matched": False,
                      "status": "未进行（HSF 001 仅支持 CAS）",
                      "records": [],
                      "source": "飞书 Wiki HSF 001 Sheet",
                      "source_note": "HSF 001 Sheet 未提供 EC 号；请输入 CAS 号进行 HSF 物质级判断。",
                      "scope_note": "仅按化合物/CAS是否命中当前HSF 001映射判断。",
                      "no_cas_count": 0,
                  })
    return {"query": value, "reach": _reach_svhc_check(value),
            "annex_xvii": _annex_xvii_check(value), "rohs": _rohs_check(value),
            "hsf_001": hsf_result}


def _rohs_list():
    """Return the complete RoHS material-level register for the read-only table page."""
    import sqlite3

    if not ROHS_DB_PATH.is_file():
        raise FileNotFoundError(f"RoHS 数据库不存在：{ROHS_DB_PATH}")
    uri = f"file:{ROHS_DB_PATH.as_posix()}?mode=ro"
    columns = ["物质类别", "中文名称", "英文名称", "CAS号", "EC号", "均质材料限值", "限值ppm", "材料筛查说明", "法规依据"]
    with sqlite3.connect(uri, uri=True) as conn:
        conn.row_factory = sqlite3.Row
        rows = [{column: row[column] or "" for column in columns} for row in conn.execute(
            'SELECT "物质类别", "中文名称", "英文名称", "CAS号", "EC号", "均质材料限值", '
            '"限值ppm", "材料筛查说明", "法规依据" FROM rohs_restricted_substance ORDER BY "序号"'
        )]
    return {"count": len(rows), "columns": columns, "rows": rows}


def _annex_xvii_check(identifier: str):
    """Read-only REACH Annex XVII CAS/EC check, including expanded group members."""
    import sqlite3

    value = str(identifier or "").strip()
    if not value:
        raise ValueError("请输入 CAS 号或 EC 号")
    if not re.fullmatch(r"\d{2,7}-\d{2,7}-\d", value):
        raise ValueError("CAS 号或 EC 号格式不正确")
    if not ANNEX_XVII_DB_PATH.is_file():
        raise FileNotFoundError(f"REACH Annex XVII 数据库不存在：{ANNEX_XVII_DB_PATH}")

    uri = f"file:{ANNEX_XVII_DB_PATH.as_posix()}?mode=ro"
    with sqlite3.connect(uri, uri=True) as conn:
        conn.row_factory = sqlite3.Row
        matched = []
        for row in conn.execute(
            'SELECT "条目号", "中文名称", "英文名称", "CAS号", "EC号", "物质描述", '
            '"限制条件标题", "限制条件链接", "法规依据", "是否物质组成员" '
            'FROM annex_xvii_substance WHERE "条目号" IN '
            '(SELECT "条目号" FROM annex_xvii_entry WHERE "中文名称状态" <> \'已删除\') '
            'ORDER BY "条目号", "记录ID"'
        ):
            cas_identifiers = re.findall(r"\d{2,7}-\d{2,7}-\d", str(row["CAS号"] or ""))
            ec_identifiers = re.findall(r"\d{2,7}-\d{2,7}-\d", str(row["EC号"] or ""))
            if value in cas_identifiers or value in ec_identifiers:
                matched.append(dict(row))
        no_identifier_entries = []
        for row in conn.execute(
            'SELECT "条目号", "中文名称", "英文名称", "CAS号", "EC号", "物质描述", '
            '"限制条件标题", "限制条件链接", "法规依据" '
            'FROM annex_xvii_entry WHERE "中文名称状态" <> \'已删除\' ORDER BY "条目号"'
        ):
            if not str(row["中文名称"] or "").strip():
                continue
            if not str(row["CAS号"] or "").strip() and not str(row["EC号"] or "").strip():
                no_identifier_entries.append(dict(row))
    return {"query": value, "matched": bool(matched), "records": matched,
            "no_identifier_entries": no_identifier_entries,
            "no_identifier_count": len(no_identifier_entries)}


def _annex_xvii_list():
    """Return active Annex XVII entries with their complete group members."""
    import sqlite3

    if not ANNEX_XVII_DB_PATH.is_file():
        raise FileNotFoundError(f"REACH Annex XVII 数据库不存在：{ANNEX_XVII_DB_PATH}")
    uri = f"file:{ANNEX_XVII_DB_PATH.as_posix()}?mode=ro"
    columns = ["条目号", "中文名称", "英文名称", "CAS号", "EC号", "物质描述", "限制条件标题", "限制条件链接", "法规依据", "物质成员"]
    with sqlite3.connect(uri, uri=True) as conn:
        conn.row_factory = sqlite3.Row
        rows = [{column: row[column] or "" for column in columns[:-1]} for row in conn.execute(
            'SELECT "条目号", "中文名称", "英文名称", "CAS号", "EC号", "物质描述", '
            '"限制条件标题", "限制条件链接", "法规依据" FROM annex_xvii_entry '
            'WHERE "中文名称状态" <> \'已删除\' '
            'ORDER BY CAST("条目号" AS INTEGER), "条目号"'
        )]
        member_sql = (
            'SELECT "记录ID", "中文名称", "英文名称", "CAS号", "EC号", "物质描述", '
            '"是否物质组成员" FROM annex_xvii_substance '
            'WHERE "条目号" = ? ORDER BY "记录ID"'
        )
        for row in rows:
            row["物质成员"] = [dict(member) for member in conn.execute(member_sql, (row["条目号"],))]
    return {"count": len(rows), "columns": columns, "rows": rows}


def _cas_s2_result(cas_no: str):
    """读取 CAS Section 2 结果骨架；结果库只读，不影响正式 CAS 库。"""
    import sqlite3

    cas = str(cas_no or "").strip()
    if not cas or not CAS_RESULT_DB_PATH.exists():
        return None


    uri = f"file:{CAS_RESULT_DB_PATH.as_posix()}?mode=ro"
    try:
        with sqlite3.connect(uri, uri=True) as conn:
            conn.row_factory = sqlite3.Row
            profile = conn.execute(
                "SELECT profile_id, cas_no, standard_name_zh, ec_no, index_no, "
                "language_code, jurisdiction, profile_version, regulatory_version, "
                "profile_status, is_current, remarks, created_at, updated_at "
                "FROM cas_s2_profile WHERE cas_no=? AND is_current=1 LIMIT 1", (cas,)
            ).fetchone()
            if not profile:
                return None
            profile_id = profile["profile_id"]
            standard_fields = [dict(r) for r in conn.execute(
                "SELECT field_key, field_label_zh, field_type, display_order, value_text, "
                "value_status FROM v_cas_s2_current_standard_result "
                "WHERE profile_id=? ORDER BY display_order", (profile_id,)
            )]
            return {
                "profile": {
                    "profile_id": profile["profile_id"],
                    "cas_no": profile["cas_no"],
                    "profile_version": profile["profile_version"],
                    "profile_status": profile["profile_status"],
                },
                "standard_fields": standard_fields,
            }
    except sqlite3.Error:
        return None


def _cas_online_search(query: str):
    """调用 CAS API 搜索程序的真实 lookup 合同，返回其原始结构。"""
    value = str(query or "").strip()
    if not value:
        raise cas_online_query.QueryError("查询内容不能为空")
    return cas_online_query.lookup(
        value,
        cas_online_query.DEFAULT_CACHE,
        cas_online_query.DEFAULT_TIMEOUT,
        True,
    )


def _library_detail(library_id: str):
    library_id = str(library_id or "").strip().lower()
    with open_db(DB_PATH) as conn:
        if library_id == "models":
            rows = conn.execute("SELECT section, COUNT(*) FROM msds_field GROUP BY section ORDER BY section").fetchall()
            return {"id": "models", "name": "中文总型号库", "metrics": {
                "型号": conn.execute("SELECT COUNT(*) FROM msds_model").fetchone()[0],
                "字段": conn.execute("SELECT COUNT(*) FROM msds_field").fetchone()[0],
                "唯一型号": conn.execute("SELECT COUNT(DISTINCT model) FROM msds_model").fetchone()[0],
                "Section": len(rows)}, "breakdown": [{"label": f"S{s}", "value": n} for s, n in rows]}
        if library_id == "field_mapping":
            rows = conn.execute("SELECT section, COUNT(*) FROM field_mapping GROUP BY section ORDER BY section").fetchall()
            return {"id": "field_mapping", "name": "字段映射库", "metrics": {
                "映射": conn.execute("SELECT COUNT(*) FROM field_mapping").fetchone()[0],
                "标准字段": conn.execute("SELECT COUNT(*) FROM schema_field").fetchone()[0],
                "待归类": conn.execute("SELECT COUNT(*) FROM msds_unmapped").fetchone()[0]},
                "breakdown": [{"label": f"S{s}", "value": n} for s, n in rows]}
    if library_id == "cas" and CAS_DB_PATH.exists():
        import sqlite3
        with sqlite3.connect(CAS_DB_PATH) as conn:
            def count(table):
                try:
                    return conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
                except sqlite3.Error:
                    return 0
            return {"id": "cas", "name": "CAS 库", "metrics": {
                "物质": count("cas_substance"),
                "确认 CAS": conn.execute("SELECT COUNT(*) FROM cas_substance WHERE TRIM(cas_no)<>''").fetchone()[0],
                "型号使用": count("cas_model_usage"), "别名": count("cas_alias")}, "breakdown": []}
    if library_id == "cas_mapping" and CAS_MAPPING_DB_PATH.exists():
        import sqlite3
        with sqlite3.connect(CAS_MAPPING_DB_PATH) as conn:
            def mapping_count(table):
                try:
                    return conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
                except sqlite3.Error:
                    return 0
            return {"id": "cas_mapping", "name": "CAS 字段映射库", "metrics": {
                "身份": mapping_count("cas_identity"),
                "字段映射": mapping_count("cas_field_mapping"),
                "观察记录": mapping_count("cas_mapping_observation"),
                "冲突": mapping_count("cas_mapping_conflict")}, "breakdown": []}
    if library_id == "categories":
        tree = _category_tree()
        return {"id": "categories", "name": "产品大类", "metrics": {
            "父级": len(tree), "型号": sum(len(x["models"]) for x in tree)}, "breakdown": [
                {"label": x["name"], "value": len(x["models"])} for x in tree]}
    return {}


def _build_form_payload(body):
    schema = build_form_schema()
    s1 = body.get("s1") or {}
    s3 = body.get("s3") or {}
    comps = [ComponentRow(str(x.get("name", "")), str(x.get("cas", "")), str(x.get("conc", "")))
             for x in (s3.get("components") or [])]
    # 固定按 37 项标准骨架排序，不能依赖浏览器对象键顺序或用户填写顺序。
    s9_values = body.get("s9") or {}
    s9 = [(f.label, str(s9_values.get(f.label, ""))) for f in schema.s9]
    return form_to_write_items(schema, s1, str(s3.get("产品类型", "混合物")), comps, s9)


def _resolve_overwrite_template(raw="", language="zh"):
    """Resolve one of the two approved language-specific overwrite templates."""
    language = str(language or "zh").lower()
    if language not in {"zh", "en"}:
        raise ValueError("语言只支持 zh 或 en")
    canonical = CANONICAL_TEMPLATE_EN if language == "en" else CANONICAL_TEMPLATE
    requested = str(raw or "").strip()
    approved = {CANONICAL_TEMPLATE.resolve(), CANONICAL_TEMPLATE_EN.resolve()}
    if requested:
        requested_path = Path(requested).resolve()
        if requested_path not in approved:
            raise ValueError("覆写只允许使用项目批准的中文或英文固定模板")
        if requested_path != canonical.resolve():
            raise ValueError(f"{language} 产出必须使用对应的固定模板：{canonical}")
    if not canonical.is_file():
        raise FileNotFoundError(f"固定{language}模板不存在：{canonical}")
    return canonical


def _form_state_snapshot(state):
    """Serialize the full 17-section console form state for the Web UI."""
    fields = []
    for key, item in state.fields.items():
        fields.append({
            "key": key, "section": item.section, "seq": item.seq,
            "label": item.label, "parent": item.parent, "current": item.current,
            "kind": item.kind,
        })
    return {
        "template": str(state.template),
        "fields": fields,
        "values": dict(state.values),
        "components": list(state.components),
        "product_type": state.product_type,
        "bio_rows": list(state.bio_rows),
        "laws": list(state.laws),
        "s15_special": dict(state.s15_special),
        "s16_special": dict(state.s16_special),
        "label_overrides": dict(state.label_overrides),
    }


def _web_form_state(body, template):
    """Build the existing Python form state from a Web request."""
    state = form_overwrite._load_state(template)
    incoming_values = body.get("values") or {}
    if not isinstance(incoming_values, dict):
        raise ValueError("values 必须是对象")
    for key, value in incoming_values.items():
        if key in state.values:
            state.values[key] = form_overwrite._clean(value)
    cleared = body.get("cleared_keys") or []
    state.cleared_keys = {str(x) for x in cleared if str(x) in state.values}
    touched = body.get("touched_sections") or []
    state.touched_sections = {int(x) for x in touched if str(x).isdigit() and 0 <= int(x) <= 16}
    components = body.get("components") or []
    state.components = [{
        "name": form_overwrite._clean(x.get("name", "")),
        "cas": form_overwrite._clean(x.get("cas", "")),
        "conc": form_overwrite._clean(x.get("conc", "")),
    } for x in components if isinstance(x, dict)]
    state.product_type = form_overwrite._clean(body.get("product_type")) or "混合物"
    bio_rows = body.get("bio_rows") or []
    state.bio_rows = [list(map(form_overwrite._clean, row[:5])) for row in bio_rows
                      if isinstance(row, list) and len(row) == 5]
    state.laws = [form_overwrite._clean(x) for x in (body.get("laws") or []) if form_overwrite._clean(x)]
    state.s15_special = {str(k): form_overwrite._clean(v) for k, v in (body.get("s15_special") or {}).items()}
    state.s16_special = {str(k): form_overwrite._clean(v) for k, v in (body.get("s16_special") or {}).items()}
    # S16 的显式别名与统一 values 同步，兼容 Web GUI/Agent 只提交特殊槽位
    # 而没有回传完整表单 values 的情况。前端既可能提交稳定字段键
    # ``S16||免责声明``，也可能只提交显示标签 ``免责声明``；两者都必须
    # 回写到 FormState.fields 对应的内部键，否则专用槽位会在预览中丢失。
    for key, value in state.s16_special.items():
        if key in state.values:
            state.values[key] = value
            continue
        for field_key, field_item in state.fields.items():
            if field_item.section == 16 and field_item.label == key:
                state.values[field_key] = value
                break
    state.label_overrides = {str(k): form_overwrite._clean(v) for k, v in (body.get("label_overrides") or {}).items()}
    return state


def _safe_web_output(name=""):
    """Return a safe DOCX/PDF output path inside the formal Web output directory."""
    raw = Path(str(name or "").strip()).name
    if not raw:
        raw = f"MSDS_交互覆写_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    if Path(raw).suffix.lower() not in {".docx", ".pdf"}:
        raw += ".docx"
    if raw in {".", ".."} or not re.fullmatch(r"[\w\-. 一-龥（）()]+\.(?:docx|pdf)", raw, re.IGNORECASE):
        raise ValueError("输出文件名只允许使用中文、字母、数字、空格、下划线、连字符和 .docx/.pdf")
    WEB_OVERWRITE_ROOT.mkdir(parents=True, exist_ok=True)
    return WEB_OVERWRITE_ROOT / raw


def _safe_multiformat_output(name="", output_format="docx"):
    """Return a safe output path for DOCX or PDF multi-format drafts."""
    fmt = str(output_format or "docx").lower().lstrip(".")
    if fmt not in {"docx", "pdf"}:
        raise ValueError("产出格式只支持 Word 或 PDF")
    raw = Path(str(name or "").strip()).name
    if not raw:
        raw = f"MSDS_多格式产出_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{fmt}"
    raw = Path(raw).stem + f".{fmt}"
    if not re.fullmatch(r"[\w\-. 一-龥（）()]+\.(?:docx|pdf)", raw, re.IGNORECASE):
        raise ValueError("输出文件名只允许使用中文、字母、数字、空格、下划线、连字符和 .docx/.pdf")
    WEB_OVERWRITE_ROOT.mkdir(parents=True, exist_ok=True)
    return WEB_OVERWRITE_ROOT / raw


def _apply_company_profile(state, company: str, language: str):
    profile = multiformat.company_profile(company, language)
    label_map = {"公司名称": "company_name", "供应商名称": "supplier_name", "供应商地址": "supplier_address", "电话": "telephone", "传真": "fax"}
    for key in list(state.values):
        label = key.split("||", 1)[-1]
        source_key = label_map.get(label)
        if source_key and source_key in profile:
            state.values[key] = profile[source_key]
    return state


def _normalize_company(company: str) -> str:
    value = str(company or "guanzhi").strip().lower()
    if value not in {"guanzhi", "guocai"}:
        raise ValueError("公司只支持 guanzhi 或 guocai")
    return value


def _safe_batch_stem(name=""):
    """Return a safe base name shared by the four Chinese batch outputs."""
    raw = Path(str(name or "").strip()).name
    raw = Path(raw).stem if raw else ""
    if not raw:
        raw = f"MSDS_中文四件套_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    if not re.fullmatch(r"[\w\-. 一-龥（）()]+", raw):
        raise ValueError("批量输出名称只允许使用中文、字母、数字、空格、下划线、连字符和圆括号")
    return raw


def _safe_batch_archive(name=""):
    """Return a safe ZIP path inside the formal Web output directory."""
    raw = Path(str(name or "").strip()).name
    raw = Path(raw).stem if raw else f"MSDS_中文四件套_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    raw = f"{raw}.zip"
    if not re.fullmatch(r"[\w\-. 一-龥（）()]+\.zip", raw, re.IGNORECASE):
        raise ValueError("批量下载包名称不合法")
    WEB_OVERWRITE_ROOT.mkdir(parents=True, exist_ok=True)
    return WEB_OVERWRITE_ROOT / raw


def _produce_overwrite(body, template, language, company, output_format, output_name=""):
    """Produce one DOCX/PDF from the shared overwrite state and return its result."""
    company = _normalize_company(company)
    model = str(body.get("model") or "")
    translation_meta = None
    if body.get("form_state") or "values" in body:
        raw_state = body.get("form_state") or body
        raw_values = raw_state.get("values") or {}
        state = _web_form_state(raw_state, template)
        if not model:
            # Prefer the submitted values so model detection still works when
            # the template normalizes or omits a field during state building.
            joined_values = " ".join(str(value or "") for value in raw_values.values())
            joined_values += " " + " ".join(str(value or "") for value in state.values.values())
            # Keep a serialized-body fallback for clients whose nested form
            # object is normalized by their JSON bridge before reaching us.
            joined_values += " " + json.dumps(body, ensure_ascii=False)
            product_match = re.search(r"\b[A-Z]{1,8}-\d+[A-Z0-9+]*\b", joined_values, re.IGNORECASE)
            model = product_match.group(0).upper() if product_match else ""
        changed_before_translation = {
            key for key, item in state.fields.items()
            if key in state.values and (
                state.values[key] != item.current or key in state.cleared_keys
            )
        }
        if language == "en":
            translated = multiformat.translate_values(dict(state.values), language, company, model)
            state.values = translated["values"]
            # 翻译会改变未修改模板值的文本；同步 EditableField.current，
            # 否则它们会被误判为 Agent 新提交的 S11/S12 证据，导致旧模板
            # 毒理/生态行重新进入正式输出。
            for key, item in state.fields.items():
                if key not in changed_before_translation and key in state.values and item.current:
                    item.current = state.values[key]
            # S15 的法规说明/条目位于独立动态容器，不在 state.values 中；
            # 同一翻译术语通道必须覆盖它们，否则 Web EN 输出会把法规正文
            # 留在中文而前端无法发现字段不匹配。
            state.s15_special = {
                key: multiformat.translate_text(value, "en")["text"]
                for key, value in state.s15_special.items()
            }
            state.laws = [multiformat.translate_text(value, "en")["text"] for value in state.laws]
            translation_meta = {k: translated[k] for k in ("review", "memory", "source_language", "target_language", "company")}
        _apply_company_profile(state, company, language)
        # Company selection is itself a Section 1 business edit: include S1
        # even when the user did not type into those fields this session.
        # S0 company identity is handled in the header/footer postprocessor;
        # adding S0 write-items would violate the existing duplicate-field
        # permission rule between S0 company name and S1 supplier name.
        state.touched_sections.add(1)
        payload = form_overwrite.build_write_items(state)
        form_overwrite.validate_write_items_permissions(payload)
    else:
        payload = body.get("payload") or _build_form_payload(body)
    output = _safe_multiformat_output(output_name or body.get("output_path", ""), output_format)
    docx_output = output if output_format == "docx" else output.with_suffix(".docx")
    logs = overwrite_engine.overwrite(
        template, payload, docx_output,
        empty_policy=payload.get("empty_policy", "overwrite"),
        missing_policy=payload.get("missing_policy", "preserve"),
        missing_text=payload.get("missing_text", ""),
        field_map=body.get("field_map") or None,
    )
    ok, problems = overwrite_engine.verify_output(
        template, docx_output, payload,
        sections=set(int(x) for x in payload.get("sections", {}) if str(x).isdigit()),
    )
    if language == "en":
        translated_doc = multiformat.translate_docx_in_place(docx_output, language, company, model)
        if translation_meta is None:
            translation_meta = {}
        translation_meta.update({"document_segments": translated_doc["segments"],
                                 "sample_count": translated_doc["sample_count"],
                                 "document_review": translated_doc["review"]})
    branding = multiformat.apply_company_branding(docx_output, company, language)
    if translation_meta is None:
        translation_meta = {}
    translation_meta["company_branding"] = branding
    if output_format == "pdf":
        from msds_multiformat import convert_docx_to_pdf
        convert_docx_to_pdf(docx_output, output)
    write_items_path = docx_output.with_suffix(".write_items.json")
    write_items_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return {
        "ok": bool(ok), "output": str(output), "output_name": output.name,
        "download_url": f"/api/overwrite/download?name={output.name}",
        "write_items": str(write_items_path), "logs": logs[-60:], "problems": problems,
        "language": language, "company": company, "output_format": output_format,
        "formal_ready": bool(ok and not (translation_meta and translation_meta.get("document_review"))),
        "translation": translation_meta,
    }


class Handler(BaseHTTPRequestHandler):
    server_version = "MSDSWeb/1.0"

    def log_message(self, fmt, *args):
        print("[web] " + fmt % args)

    def send_json(self, payload, status=HTTPStatus.OK):
        data = json.dumps(payload, ensure_ascii=False, default=_jsonable).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def send_file(self, path):
        data = path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", mimetypes.guess_type(str(path))[0] or "application/octet-stream")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate")
        self.end_headers()
        self.wfile.write(data)

    def send_download(self, path):
        data = path.read_bytes()
        suffix = path.suffix.lower()
        content_type = mimetypes.guess_type(str(path))[0] or "application/octet-stream"
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Content-Disposition", f"attachment; filename=msds-output{suffix}; filename*=UTF-8''{quote(path.name)}")
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def read_body(self):
        length = int(self.headers.get("Content-Length", "0"))
        return json.loads(self.rfile.read(length) or b"{}")

    def read_upload(self):
        """Read one .docx from a multipart request without adding a dependency."""
        from email.parser import BytesParser
        from email.policy import default

        content_type = self.headers.get("Content-Type", "")
        if not content_type.lower().startswith("multipart/form-data"):
            raise ValueError("上传请求必须是 multipart/form-data")
        length = int(self.headers.get("Content-Length", "0"))
        raw = self.rfile.read(length)
        message = BytesParser(policy=default).parsebytes(
            (f"Content-Type: {content_type}\r\n"
             "MIME-Version: 1.0\r\n\r\n").encode() + raw)
        for part in message.walk():
            filename = part.get_filename()
            if filename:
                return filename, part.get_payload(decode=True) or b""
        raise ValueError("未找到上传文件")

    def import_upload(self):
        filename, content = self.read_upload()
        name = Path(filename).name
        if not name.lower().endswith(".docx"):
            return self.send_json({"error": "仅支持 .docx 格式的 MSDS Word 文档"}, HTTPStatus.BAD_REQUEST)
        if len(content) > 20 * 1024 * 1024:
            return self.send_json({"error": "文件超过 20MB 限制"}, HTTPStatus.REQUEST_ENTITY_TOO_LARGE)
        tmpdir = Path(tempfile.mkdtemp(prefix="msds_import_"))
        tmp = tmpdir / name
        try:
            tmp.write_bytes(content)
            result = read_msds(tmp)
            sections = build_hierarchy(result)
            raw_by_num = {int(section.number): section for section in sections}
            # GUI 主窗口不是直接渲染 build_hierarchy 原始节点，而是统一经过
            # listed_rows_from_result() 生成 S0~S16 标准骨架。Web 导入必须走同一条链路，
            # 否则空字段、父级标题、通栏行和子表结构会在前端 flatten 时丢失。
            standard_sections = []
            for number in range(17):
                raw = raw_by_num.get(number)
                title = getattr(raw, "title", f"第{number}节") if raw else f"第{number}节"
                full_title = getattr(raw, "full_title", title) if raw else title
                standard_sections.append({
                    "number": number,
                    "title": title,
                    "full_title": full_title,
                    "rows": [_jsonable(row) for row in listed_rows_from_result(result, number)],
                })
            return self.send_json({
                "ok": True,
                "filename": name,
                "model": getattr(result, "model", ""),
                "sections": standard_sections,
            })
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

    def do_GET(self):
        parsed = urlparse(self.path)
        try:
            if parsed.path == "/api/health":
                with open_db(DB_PATH) as conn:
                    count = conn.execute("SELECT COUNT(*) FROM msds_model").fetchone()[0]
                return self.send_json({"ok": True, "models": count, "db": str(DB_PATH)})
            if parsed.path == "/api/overwrite/templates":
                if not CANONICAL_TEMPLATE.is_file() or not CANONICAL_TEMPLATE_EN.is_file():
                    missing = [str(x) for x in (CANONICAL_TEMPLATE, CANONICAL_TEMPLATE_EN) if not x.is_file()]
                    raise FileNotFoundError(f"固定模板不存在：{', '.join(missing)}")
                return self.send_json({"default": str(CANONICAL_TEMPLATE), "templates": [
                    {"name": CANONICAL_TEMPLATE.name, "path": str(CANONICAL_TEMPLATE), "language": "zh", "approved": True},
                    {"name": CANONICAL_TEMPLATE_EN.name, "path": str(CANONICAL_TEMPLATE_EN), "language": "en", "approved": True},
                ]})
            if parsed.path == "/api/overwrite/form":
                query = parse_qs(parsed.query)
                requested = query.get("template", [""])[0]
                language = query.get("language", [""])[0]
                if not language and requested:
                    language = "en" if Path(requested).resolve() == CANONICAL_TEMPLATE_EN.resolve() else "zh"
                template = _resolve_overwrite_template(requested, language or "zh")
                company = _normalize_company(query.get("company", ["guanzhi"])[0])
                state = form_overwrite._load_state(template)
                _apply_company_profile(state, company, language or "zh")
                return self.send_json({"ok": True, "company": company,
                                       "form": _form_state_snapshot(state)})
            if parsed.path == "/api/multiformat/catalog":
                return self.send_json({"ok": True, "companies": multiformat.profile_options(),
                                       "languages": [{"id": "zh", "label": "中文"}, {"id": "en", "label": "English"}],
                                       "formats": [{"id": "docx", "label": "Word"}, {"id": "pdf", "label": "PDF"}]})
            if parsed.path == "/api/tds/catalog":
                query = parse_qs(parsed.query)
                return self.send_json(tds_web.catalog(
                    query.get("q", [""])[0], query.get("language", [""])[0],
                    query.get("company", [""])[0], query.get("category", [""])[0],
                    int(query.get("limit", [250])[0] or 250)))
            if parsed.path == "/api/tds/templates":
                return self.send_json(tds_web.templates())
            if parsed.path == "/api/tds/source":
                query = parse_qs(parsed.query)
                source_id = query.get("id", [""])[0]
                language = query.get("language", ["zh"])[0]
                company = query.get("company", ["guanzhi"])[0]
                return self.send_json({"ok": True, "form": tds_web.load_source(source_id, language, company)})
            if parsed.path == "/api/tds/download":
                name = Path(parse_qs(parsed.query).get("name", [""])[0]).name
                return self.send_download(tds_web.safe_output(name))
            if parsed.path == "/api/tds/batch-download":
                name = Path(parse_qs(parsed.query).get("name", [""])[0]).name
                return self.send_download(tds_web.safe_output(name))
            if parsed.path == "/api/ghs/pictograms":
                return self.send_json({"ok": True, "items": _ghs_pictogram_catalog()})
            if parsed.path == "/api/overwrite/download":
                name = Path(parse_qs(parsed.query).get("name", [""])[0]).name
                output = _safe_web_output(name)
                if not output.is_file():
                    return self.send_json({"error": "产出文件不存在"}, HTTPStatus.NOT_FOUND)
                return self.send_download(output)
            if parsed.path == "/api/overwrite/batch-download":
                name = Path(parse_qs(parsed.query).get("name", [""])[0]).name
                archive = _safe_batch_archive(name)
                if not archive.is_file():
                    return self.send_json({"error": "批量下载包不存在"}, HTTPStatus.NOT_FOUND)
                return self.send_download(archive)
            if parsed.path == "/api/libraries":
                return self.send_json({"items": _library_info()})
            if parsed.path == "/api/categories":
                return self.send_json({"items": _category_tree()})
            if parsed.path == "/api/cas":
                q = parse_qs(parsed.query).get("q", [""])[0]
                status = parse_qs(parsed.query).get("status", [""])[0]
                if not CAS_DB_PATH.exists():
                    return self.send_json({"items": [], "error": "CAS 库不存在"}, HTTPStatus.NOT_FOUND)
                return self.send_json({"items": _cas_search(q, status), "query": q, "status": status})
            if parsed.path == "/api/cas-online":
                q = parse_qs(parsed.query).get("q", [""])[0]
                try:
                    return self.send_json({"ok": True, "result": _cas_online_search(q)})
                except cas_online_query.QueryError as exc:
                    return self.send_json({"ok": False, "query": q, "error": str(exc)}, HTTPStatus.NOT_FOUND)
            if parsed.path == "/api/reach/check":
                params = parse_qs(parsed.query)
                q = params.get("identifier", params.get("cas", [""]))[0]
                try:
                    return self.send_json({"ok": True, "result": _reach_svhc_check(q)})
                except ValueError as exc:
                    return self.send_json({"ok": False, "query": q, "error": str(exc)}, HTTPStatus.BAD_REQUEST)
                except FileNotFoundError as exc:
                    return self.send_json({"ok": False, "query": q, "error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)
            if parsed.path == "/api/reach/list":
                try:
                    return self.send_json({"ok": True, "result": _reach_svhc_list()})
                except FileNotFoundError as exc:
                    return self.send_json({"ok": False, "error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)
            if parsed.path == "/api/rohs/check":
                params = parse_qs(parsed.query)
                q = params.get("identifier", params.get("cas", [""]))[0]
                try:
                    return self.send_json({"ok": True, "result": _rohs_check(q)})
                except ValueError as exc:
                    return self.send_json({"ok": False, "query": q, "error": str(exc)}, HTTPStatus.BAD_REQUEST)
                except FileNotFoundError as exc:
                    return self.send_json({"ok": False, "query": q, "error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)
            if parsed.path == "/api/rohs/list":
                try:
                    return self.send_json({"ok": True, "result": _rohs_list()})
                except FileNotFoundError as exc:
                    return self.send_json({"ok": False, "error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)
            if parsed.path == "/api/annex-xvii/check":
                params = parse_qs(parsed.query)
                q = params.get("identifier", params.get("cas", [""]))[0]
                try:
                    return self.send_json({"ok": True, "result": _annex_xvii_check(q)})
                except ValueError as exc:
                    return self.send_json({"ok": False, "query": q, "error": str(exc)}, HTTPStatus.BAD_REQUEST)
                except FileNotFoundError as exc:
                    return self.send_json({"ok": False, "query": q, "error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)
            if parsed.path == "/api/annex-xvii/list":
                try:
                    return self.send_json({"ok": True, "result": _annex_xvii_list()})
                except FileNotFoundError as exc:
                    return self.send_json({"ok": False, "error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)
            if parsed.path == "/api/hsf-001/check":
                params = parse_qs(parsed.query)
                q = params.get("identifier", params.get("cas", [""]))[0]
                try:
                    return self.send_json({"ok": True, "result": _hsf_001_check(q)})
                except ValueError as exc:
                    return self.send_json({"ok": False, "query": q, "error": str(exc)}, HTTPStatus.BAD_REQUEST)
                except FileNotFoundError as exc:
                    return self.send_json({"ok": False, "query": q, "error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)
            if parsed.path == "/api/hsf-001/list":
                try:
                    return self.send_json({"ok": True, "result": _hsf_001_list()})
                except FileNotFoundError as exc:
                    return self.send_json({"ok": False, "error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)
            if parsed.path == "/api/compliance/check":
                params = parse_qs(parsed.query)
                q = params.get("identifier", params.get("cas", [""]))[0]
                try:
                    return self.send_json({"ok": True, "result": _compliance_check(q)})
                except ValueError as exc:
                    return self.send_json({"ok": False, "query": q, "error": str(exc)}, HTTPStatus.BAD_REQUEST)
                except FileNotFoundError as exc:
                    return self.send_json({"ok": False, "query": q, "error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)
            if parsed.path.startswith("/api/libraries/"):
                detail = _library_detail(parsed.path.rsplit("/", 1)[-1])
                if detail:
                    return self.send_json(detail)
                return self.send_json({"error": "library not found"}, HTTPStatus.NOT_FOUND)
            if parsed.path == "/api/schema/form":
                return self.send_json(_schema())
            if parsed.path == "/api/models":
                q = parse_qs(parsed.query).get("q", [""])[0]
                category = parse_qs(parsed.query).get("category", [""])[0]
                with open_db(DB_PATH) as conn:
                    if q:
                        like = f"%{q.strip()}%"
                        rows = conn.execute(
                            "SELECT DISTINCT m.model_id, m.model, m.source, m.source_file, mc.category_parent, "
                            "(SELECT COUNT(*) FROM msds_field f2 WHERE f2.model_id=m.model_id), m.created_at "
                            ",(SELECT GROUP_CONCAT(fh.label || ': ' || substr(COALESCE(fh.value,''), 1, 80), '；') "
                            "FROM msds_field fh WHERE fh.model_id=m.model_id "
                            "AND (fh.label LIKE ? OR fh.std_name LIKE ? OR fh.value LIKE ?)) AS hits "
                            "FROM msds_model m LEFT JOIN msds_field f ON f.model_id=m.model_id "
                            "LEFT JOIN model_category mc ON mc.model_id=m.model_id "
                            "WHERE (m.model LIKE ? OR f.label LIKE ? OR f.std_name LIKE ? OR f.value LIKE ?) "
                            "AND (?='' OR mc.category_parent=?) "
                            "ORDER BY CASE WHEN m.model=? THEN 0 ELSE 1 END, m.model",
                            (like, like, like, like, like, like, like, category, category, q.strip())).fetchall()
                    else:
                        rows = conn.execute(
                            "SELECT m.model_id, m.model, m.source, m.source_file, mc.category_parent, "
                            "(SELECT COUNT(*) FROM msds_field f2 WHERE f2.model_id=m.model_id), m.created_at "
                            "FROM msds_model m LEFT JOIN model_category mc ON mc.model_id=m.model_id "
                            "WHERE (?='' OR mc.category_parent=?) "
                            "ORDER BY m.model", (category, category)).fetchall()
                return self.send_json({"items": [
                     {"id": r[0], "model_id": r[0], "model": r[1], "source": r[2], "source_file": r[3],
                      "category": r[4] if len(r) > 6 else "", "fields": r[5] if len(r) > 6 else r[4],
                      "fields_count": r[5] if len(r) > 6 else r[4],
                      "created_at": r[6] if len(r) > 6 else r[5],
                      "hits": r[7] if len(r) > 7 else ""} for r in rows]})
            parts = parsed.path.strip("/").split("/")
            if len(parts) >= 3 and parts[0] == "api" and parts[1] == "models":
                model_id = int(parts[2])
                with open_db(DB_PATH) as conn:
                    if len(parts) == 3:
                        return self.send_json({"detail": model_detail(conn, model_id),
                                               "wide": wide_row(conn, model_id),
                                               "rows": _rows(conn, model_id)})
                    if parts[3] == "form":
                        return self.send_json(_form_from_model(conn, model_id))
                    if parts[3] == "write-items":
                        return self.send_json(model_to_write_items(conn, model_id, s9_active_only=False))
            if parsed.path.startswith("/assets/"):
                path = (WEB_ROOT / parsed.path.removeprefix("/"))
                if path.is_file() and WEB_ROOT in path.resolve().parents:
                    return self.send_file(path)
            path = WEB_ROOT / "index.html" if parsed.path in ("/", "") else WEB_ROOT / parsed.path.removeprefix("/")
            if path.is_file() and WEB_ROOT in path.resolve().parents:
                return self.send_file(path)
            return self.send_json({"error": "not found"}, HTTPStatus.NOT_FOUND)
        except Exception as exc:
            traceback.print_exc()
            return self.send_json({"error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)

    def do_POST(self):
        parsed = urlparse(self.path)
        try:
            if parsed.path == "/api/import/upload":
                return self.import_upload()
            body = self.read_body()
            if parsed.path == "/api/form/preview":
                return self.send_json({"payload": _build_form_payload(body)})
            if parsed.path == "/api/multiformat/translate-preview":
                language = str(body.get("language") or "en").lower()
                company = str(body.get("company") or "guanzhi").lower()
                result = multiformat.translate_values(body.get("values") or {}, language, company, body.get("model", ""))
                result["company_profile"] = multiformat.company_profile(company, language)
                return self.send_json({"ok": True, "result": result})
            if parsed.path == "/api/tds/preview":
                return self.send_json(tds_web.preview(
                    str(body.get("source_id") or ""), str(body.get("language") or "zh"),
                    str(body.get("company") or "guanzhi"), body.get("form") or None))
            if parsed.path == "/api/tds/generate":
                return self.send_json(tds_web.generate(
                    str(body.get("source_id") or ""), str(body.get("language") or "zh"),
                    str(body.get("company") or "guanzhi"), str(body.get("output_format") or "docx"),
                    body.get("form") or None, str(body.get("output_name") or ""),
                    str(body.get("form_language") or body.get("language") or "zh")))
            if parsed.path == "/api/tds/batch":
                return self.send_json(tds_web.batch(
                    str(body.get("source_id") or ""), body.get("form") or None,
                    str(body.get("output_name") or ""),
                    str(body.get("form_language") or "zh")))
            if parsed.path == "/api/overwrite/preview":
                language = str(body.get("language") or "zh").lower()
                company = _normalize_company(body.get("company") or "guanzhi")
                template = _resolve_overwrite_template(body.get("template_path", ""), language)
                state = _web_form_state(body.get("form_state") or body, template)
                _apply_company_profile(state, company, language)
                state.touched_sections.add(1)
                payload = form_overwrite.build_write_items(state)
                form_overwrite.validate_write_items_permissions(payload)
                sections = payload.get("sections", {})
                summary = {str(sec): (len(items) if isinstance(items, list) else len(items.get("components", [])))
                           for sec, items in sections.items()}
                return self.send_json({"ok": True, "payload": payload, "summary": summary})
            if parsed.path.startswith("/api/drafts/"):
                model_id = parsed.path.rsplit("/", 1)[-1]
                DRAFT_ROOT.mkdir(parents=True, exist_ok=True)
                (DRAFT_ROOT / f"{model_id}.json").write_text(json.dumps(body, ensure_ascii=False, indent=2), encoding="utf-8")
                return self.send_json({"ok": True, "saved": model_id})
            if parsed.path == "/api/overwrite/batch":
                language = str(body.get("language") or "zh").lower()
                if language != "zh":
                    return self.send_json({"error": "中文四件套产出只支持中文表单，请先切换到中文"}, HTTPStatus.BAD_REQUEST)
                template = _resolve_overwrite_template(body.get("template_path", ""), "zh")
                base = _safe_batch_stem(body.get("output_name") or body.get("model", ""))
                company_names = {"guanzhi": "冠志", "guocai": "国彩"}
                format_names = {"docx": "Word", "pdf": "PDF"}
                results = []
                for company in ("guanzhi", "guocai"):
                    for output_format in ("docx", "pdf"):
                        output_name = f"{base}_{company_names[company]}_{format_names[output_format]}.{output_format}"
                        try:
                            result = _produce_overwrite(body, template, "zh", company, output_format, output_name)
                            results.append(result)
                        except Exception as exc:
                            traceback.print_exc()
                            results.append({
                                "ok": False, "output_name": output_name, "company": company,
                                "output_format": output_format, "error": str(exc), "problems": [],
                            })
                successful = [item for item in results if item.get("ok") and Path(item.get("output", "")).is_file()]
                all_ok = len(successful) == 4 and len(results) == 4
                archive = None
                archive_error = ""
                if all_ok:
                    try:
                        archive = _safe_batch_archive(f"{base}_中文MSDS四件套.zip")
                        with zipfile.ZipFile(archive, "w", compression=zipfile.ZIP_DEFLATED) as bundle:
                            for item in successful:
                                bundle.write(item["output"], arcname=Path(item["output"]).name)
                    except Exception as exc:
                        traceback.print_exc()
                        archive_error = str(exc)
                        all_ok = False
                return self.send_json({
                    "ok": all_ok, "language": "zh", "outputs": results,
                    "archive_name": archive.name if archive else "",
                    "download_url": f"/api/overwrite/batch-download?name={archive.name}" if archive else "",
                    "error": archive_error or ("部分文件产出失败" if not all_ok else ""),
                })
            if parsed.path == "/api/overwrite":
                language = str(body.get("language") or "zh").lower()
                template = _resolve_overwrite_template(body.get("template_path", ""), language)
                company = _normalize_company(body.get("company") or "guanzhi")
                output_format = str(body.get("output_format") or "docx").lower()
                result = _produce_overwrite(
                    body, template, language, company, output_format,
                    body.get("output_name") or body.get("output_path", ""),
                )
                return self.send_json(result)
            return self.send_json({"error": "not found"}, HTTPStatus.NOT_FOUND)
        except Exception as exc:
            traceback.print_exc()
            return self.send_json({"error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)


def main():
    import argparse
    parser = argparse.ArgumentParser(description="MSDS 正式版 Web 程序")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    args = parser.parse_args()
    print(f"MSDS Web: http://{args.host}:{args.port}")
    ThreadingHTTPServer((args.host, args.port), Handler).serve_forever()


if __name__ == "__main__":
    main()
