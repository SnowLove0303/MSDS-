# -*- coding: utf-8 -*-
"""逐型号核查中文源文件与 SQLite 标准库的 S1/S3/S9 数据一致性。

用法:
  python audit_db_consistency.py <db> <源目录...> --out <报告.json>

比较原则:
  - 型号由源文件 S1/S0 产品名称提取，数据库按 model 唯一匹配；
  - S1/S9 按标准字段名 + 值比较，未匹配字段按原始标签 + 值比较；
  - S3 产品类型按字段比较，成分按名称/CAS/含量比较；
  - 同型号多个源文件全部参与核查，任何一个源文件不一致都标记为 fail，
    不允许用一个文件静默覆盖另一个文件。
"""
from __future__ import annotations

import argparse
import json
import re
import sqlite3
import sys
from collections import Counter, defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from core.docx_reader import read_msds
from core.msds_db import _model_of
from core.schema import standard_fields, standard_name


def norm(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip())


def model_rows(conn: sqlite3.Connection, model_id: int, section: int):
    return conn.execute(
        "SELECT label, std_name, value, kind FROM msds_field "
        "WHERE model_id=? AND section=? ORDER BY row_index",
        (model_id, section)).fetchall()


def source_model(result, file_name: str) -> str:
    return _model_of(result, file_name)


def field_counter(result, section: int) -> Counter:
    listed = {f.name for f in standard_fields(section)}
    out = Counter()
    sec = result.sections.get(section)
    if not sec:
        return out
    for row in sec.iter_rows():
        if row.kind != "field":
            continue
        label = norm(row.label)
        value = norm(row.value)
        if not label and not value:
            continue
        std = standard_name(section, label)
        key = ("mapped", std, value) if std in listed else ("unmapped", label, value)
        out[key] += 1
    return out


def db_field_counter(conn: sqlite3.Connection, model_id: int, section: int) -> Counter:
    listed = {f.name for f in standard_fields(section)}
    out = Counter()
    for label, std, value, kind in model_rows(conn, model_id, section):
        if kind != "field":
            continue
        label, std, value = norm(label), norm(std), norm(value)
        if not label and not value:
            continue
        if std in listed:
            out[("mapped", std, value)] += 1
    for label, value in conn.execute(
        "SELECT raw_label, raw_value FROM msds_unmapped "
        "WHERE model_id=? AND section=?",
        (model_id, section)):
        label, value = norm(label), norm(value)
        if label or value:
            out[("unmapped", label, value)] += 1
    return out


def component_counter(result) -> Counter:
    sec = result.sections.get(3)
    out = Counter()
    if not sec:
        return out
    for c in sec.components:
        out[(norm(c.name), norm(c.cas), norm(c.conc))] += 1
    return out


def db_component_counter(conn: sqlite3.Connection, model_id: int) -> Counter:
    out = Counter()
    for label, value in conn.execute(
        "SELECT label, value FROM msds_field "
        "WHERE model_id=? AND section=3 AND kind='component' ORDER BY row_index",
        (model_id,)):
        cas = conc = ""
        for part in norm(value).split(" | "):
            if part.startswith("CAS:"):
                cas = norm(part[4:])
            elif part.startswith("含量:"):
                conc = norm(part[3:])
        out[(norm(label), cas, conc)] += 1
    return out


def diff_counter(expected: Counter, actual: Counter) -> dict:
    missing = list((expected - actual).elements())
    extra = list((actual - expected).elements())
    return {"missing": missing, "extra": extra, "ok": not missing and not extra}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("db")
    ap.add_argument("sources", nargs="+")
    ap.add_argument("--out", required=True)
    args = ap.parse_args()
    conn = sqlite3.connect(args.db)
    files = []
    for raw in args.sources:
        p = Path(raw)
        if p.is_dir():
            files.extend(sorted(p.rglob("*.docx")))
        elif p.suffix.lower() == ".docx":
            files.append(p)
    reports = []
    for path in sorted(set(files), key=lambda x: str(x).lower()):
        try:
            result = read_msds(str(path))
            model = source_model(result, path.name)
            hit = conn.execute(
                "SELECT model_id, source_file FROM msds_model WHERE model=?", (model,)
            ).fetchone()
            item = {"model": model, "source_file": str(path), "status": "fail", "sections": {}}
            if not hit:
                item["reason"] = "数据库不存在该型号"
                reports.append(item)
                continue
            model_id, db_source = hit
            item["db_source_file"] = db_source
            item["sections"]["1"] = diff_counter(
                field_counter(result, 1), db_field_counter(conn, model_id, 1))
            expected_s3 = field_counter(result, 3)
            actual_s3 = db_field_counter(conn, model_id, 3)
            s3 = diff_counter(expected_s3, actual_s3)
            s3["components"] = diff_counter(
                component_counter(result), db_component_counter(conn, model_id))
            s3["ok"] = s3["ok"] and s3["components"]["ok"]
            item["sections"]["3"] = s3
            item["sections"]["9"] = diff_counter(
                field_counter(result, 9), db_field_counter(conn, model_id, 9))
            item["status"] = "ok" if all(
                x.get("ok") for x in item["sections"].values()) else "fail"
        except Exception as exc:
            item = {"source_file": str(path), "status": "error",
                    "reason": f"{type(exc).__name__}: {exc}"}
        reports.append(item)

    by_model = defaultdict(list)
    for item in reports:
        by_model[item.get("model", "").strip()].append(item)
    duplicate_models = {
        model: [x["source_file"] for x in items]
        for model, items in by_model.items() if model and len(items) > 1
    }
    summary = {
        "source_files": len(files),
        "unique_source_models": len([x for x in by_model if x]),
        "duplicate_source_models": duplicate_models,
        "ok": sum(x["status"] == "ok" for x in reports),
        "fail": sum(x["status"] == "fail" for x in reports),
        "error": sum(x["status"] == "error" for x in reports),
    }
    payload = {"summary": summary, "reports": reports}
    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    Path(args.out).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if summary["fail"] == 0 and summary["error"] == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
