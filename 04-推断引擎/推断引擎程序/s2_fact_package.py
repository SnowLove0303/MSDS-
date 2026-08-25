"""Read-only Section 2 fact package collector.

The collector assembles S1/S3/S9 facts from the formal Chinese model database,
links the model to the formal CAS database, and records the registered legal
sources. The raw snapshot is read-only; the separate ``inference`` node runs
the auditable Section 2 rule functions and never writes to either formal DB.
"""

from __future__ import annotations

import argparse
import json
import sqlite3
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

ENGINE_DIR = Path(__file__).resolve().parent
if str(ENGINE_DIR) not in sys.path:
    sys.path.insert(0, str(ENGINE_DIR))
from s2_inference import infer_s2  # noqa: E402


MODEL_DB = Path(r"F:\正式项目与模块化内容\冠志\MSDS\03-数据库\正式库\Data Base\msds_standard.db")
CAS_DB = Path(r"F:\正式项目与模块化内容\冠志\MSDS\03-数据库\正式库\Data Base\cas_library.db")
SOURCE_MANIFEST = ENGINE_DIR / "s2_sources.json"


def _read_only_connection(path: Path) -> sqlite3.Connection:
    if not path.exists():
        raise FileNotFoundError(f"数据库不存在: {path}")
    uri = f"file:{path.as_posix()}?mode=ro"
    connection = sqlite3.connect(uri, uri=True)
    connection.row_factory = sqlite3.Row
    return connection


def _rows(connection: sqlite3.Connection, sql: str, parameters: Iterable[Any] = ()) -> list[dict[str, Any]]:
    return [dict(row) for row in connection.execute(sql, tuple(parameters)).fetchall()]


def _value(row: dict[str, Any]) -> Any:
    value = row.get("value")
    if value is None or value == "":
        return "无数据"
    return value


def _field_name(row: dict[str, Any]) -> str:
    return str(row.get("std_name") or row.get("label") or row.get("field_name") or "未命名字段")


def _append_field(target: dict[str, Any], name: str, value: Any) -> None:
    value = "无数据" if value is None or value == "" else value
    if name not in target:
        target[name] = value
        return
    if not isinstance(target[name], list):
        target[name] = [target[name]]
    target[name].append(value)


def _collect_s1(rows: list[dict[str, Any]]) -> dict[str, Any]:
    facts: dict[str, Any] = {}
    evidence: list[dict[str, Any]] = []
    for row in rows:
        name = _field_name(row)
        value = _value(row)
        _append_field(facts, name, value)
        evidence.append(
            {
                "section": 1,
                "field": name,
                "raw_label": row.get("label"),
                "value": value,
                "source_file": row.get("source_file") or "无数据",
                "source_page": row.get("source_page") or "无数据",
            }
        )
    return {"facts": facts, "rows": rows, "evidence": evidence}


def _collect_s3(rows: list[dict[str, Any]]) -> dict[str, Any]:
    facts: dict[str, Any] = {}
    evidence: list[dict[str, Any]] = []
    for row in rows:
        name = _field_name(row)
        value = _value(row)
        _append_field(facts, name, value)
        evidence.append(
            {
                "section": 3,
                "field": name,
                "raw_label": row.get("label"),
                "value": value,
                "kind": row.get("kind") or "无数据",
                "source_file": row.get("source_file") or "无数据",
                "source_page": row.get("source_page") or "无数据",
            }
        )
    return {"facts": facts, "rows": rows, "evidence": evidence}


def _collect_s9(rows: list[dict[str, Any]]) -> dict[str, Any]:
    facts: dict[str, Any] = {}
    evidence: list[dict[str, Any]] = []
    for row in rows:
        name = _field_name(row)
        value = _value(row)
        _append_field(facts, name, value)
        evidence.append(
            {
                "section": 9,
                "field": name,
                "raw_label": row.get("label"),
                "value": value,
                "unit": row.get("unit") or "无数据",
                "source_file": row.get("source_file") or "无数据",
                "source_page": row.get("source_page") or "无数据",
            }
        )
    return {"facts": facts, "rows": rows, "evidence": evidence}


def _collect_s2_schema(connection: sqlite3.Connection) -> dict[str, Any]:
    schema_rows = _rows(
        connection,
        "SELECT * FROM schema_field WHERE section = 2 ORDER BY id",
    )
    if not schema_rows:
        raise RuntimeError("正式库 schema_field 未找到 Section 2 骨架")
    mapping_rows = _rows(
        connection,
        """
        SELECT * FROM field_mapping
        WHERE section = 2
        ORDER BY standard_name, mapping_id
        """,
    )
    package = {
        "rows": schema_rows,
        "editable_fields": [
            row["name"] for row in schema_rows if row.get("kind") == "field"
        ],
        "group_rows": [
            row["name"] for row in schema_rows if row.get("kind") == "sub"
        ],
        "field_mapping_rows": mapping_rows,
    }
    return package


def _collect_cas(model: str, connection: sqlite3.Connection) -> dict[str, Any]:
    usage = _rows(
        connection,
        """
        SELECT u.usage_id, u.cas_id, u.model, u.raw_name, u.raw_cas,
               u.concentration, u.source_file,
               s.cas_no, s.standard_name, s.category, s.registry_status,
               s.source_count, s.model_count, s.remarks
        FROM cas_model_usage AS u
        LEFT JOIN cas_substance AS s ON s.cas_id = u.cas_id
        WHERE u.model = ?
        ORDER BY u.usage_id
        """,
        [model],
    )
    aliases: dict[int, list[dict[str, Any]]] = {}
    for row in usage:
        cas_id = row.get("cas_id")
        if cas_id is None:
            continue
        aliases[cas_id] = _rows(
            connection,
            "SELECT alias, alias_type, occurrence_count FROM cas_alias WHERE cas_id = ? ORDER BY alias_id",
            [cas_id],
        )
    for row in usage:
        row["aliases"] = aliases.get(row.get("cas_id"), [])
    package = {
        "model": model,
        "usage": usage,
        "substance_count": len({row.get("cas_id") for row in usage if row.get("cas_id") is not None}),
        "usage_count": len(usage),
    }
    return package


def _load_model(model: str, connection: sqlite3.Connection) -> dict[str, Any]:
    model_rows = _rows(connection, "SELECT * FROM msds_model WHERE model = ?", [model])
    if not model_rows:
        raise LookupError(f"正式中文总型号库未找到型号: {model}")
    if len(model_rows) > 1:
        raise RuntimeError(f"正式中文总型号库存在重复型号，已停止采集: {model}")
    model_row = model_rows[0]
    section_rows = _rows(
        connection,
        """
        SELECT * FROM msds_field
        WHERE model_id = ? AND section IN (1, 3, 9)
        ORDER BY section, id
        """,
        [model_row["model_id"]],
    )
    by_section: dict[int, list[dict[str, Any]]] = {1: [], 3: [], 9: []}
    for row in section_rows:
        by_section.setdefault(int(row["section"]), []).append(row)
    return {
        "model": model_row,
        "s1": _collect_s1(by_section[1]),
        "s3": _collect_s3(by_section[3]),
        "s9": _collect_s9(by_section[9]),
    }


def collect_fact_package(model: str) -> dict[str, Any]:
    manifest = json.loads(SOURCE_MANIFEST.read_text(encoding="utf-8"))
    with _read_only_connection(MODEL_DB) as model_connection:
        model_snapshot = _load_model(model, model_connection)
        s2_schema = _collect_s2_schema(model_connection)
    with _read_only_connection(CAS_DB) as cas_connection:
        cas_snapshot = _collect_cas(model, cas_connection)

    source_ids = [
        source["id"]
        for source_group in ("input_databases", "structured_sources", "legal_sources")
        for source in manifest.get(source_group, [])
    ]
    package = {
        "package_version": "0.1.0",
        "engine": "section2",
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "model": model,
        "status": "manual_review",
        "write_formal_database": False,
        "input_snapshot": {
            "model_record": model_snapshot["model"],
            "s1": model_snapshot["s1"],
            "s3": model_snapshot["s3"],
            "s9": model_snapshot["s9"],
            "s2_schema": s2_schema,
        },
        "substance_snapshot": cas_snapshot,
        "applicable_regulations": [],
        "candidate_classification": [],
        "candidate_label_elements": {
            "pictograms": [],
            "signal_word": "无数据",
            "hazard_statements": [],
            "precautionary_statements": [],
        },
        "s2_fields": {name: "无数据" for name in s2_schema["editable_fields"]},
        "evidence": [],
        "registered_source_ids": source_ids,
        "source_gaps": manifest.get("source_gaps", []),
        "review_notes": [
            "input_snapshot 只收集 S1/S3/S9 和 CAS 原始事实；inference 节单独保存规则计算结果。",
            "法规命中、分类、H/P 代码和象形图只能由 s2_inference 的可审计条件生成。",
            "空值按系统约定保留为“无数据”，并同时保留原始行和原始来源信息。",
        ],
    }
    package["inference"] = infer_s2(package)
    return package


def main() -> int:
    parser = argparse.ArgumentParser(description="只读生成 Section 2 事实包")
    parser.add_argument("--model", required=True, help="正式中文总型号库中的唯一型号")
    parser.add_argument("--out", type=Path, help="输出 JSON 文件路径；不传则输出到标准输出")
    args = parser.parse_args()
    package = collect_fact_package(args.model)
    serialized = json.dumps(package, ensure_ascii=False, indent=2)
    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(serialized + "\n", encoding="utf-8")
    else:
        print(serialized)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
