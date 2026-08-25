# -*- coding: utf-8 -*-
"""验证 CAS 身份库、型号关联和独立 CAS 字段映射库。"""
from __future__ import annotations

import argparse
import json
import sqlite3
from pathlib import Path


def _ro(path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(f"file:{path.as_posix()}?mode=ro", uri=True)
    conn.row_factory = sqlite3.Row
    return conn


def _cas_checksum_ok(value: str) -> bool:
    parts = value.split("-")
    if len(parts) != 3:
        return False
    digits = "".join(parts)
    if not digits.isdigit() or len(digits) < 3:
        return False
    total = sum(int(char) * weight for weight, char in enumerate(reversed(digits[:-1]), 1))
    return total % 10 == int(digits[-1])


def validate(cas_db: Path, mapping_db: Path | None = None,
             model_db: Path | None = None) -> dict:
    errors: list[str] = []
    warnings: list[str] = []
    cas = _ro(cas_db)
    tables = {row["name"] for row in cas.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")}
    expected = {"cas_substance", "cas_alias", "cas_model_usage", "cas_note"}
    missing_tables = sorted(expected - tables)
    if missing_tables:
        errors.append(f"正式 CAS 库缺少表: {missing_tables}")

    substances = cas.execute("SELECT * FROM cas_substance").fetchall()
    valid_rows = [row for row in substances if str(row["cas_no"] or "").strip()]
    duplicate_cas = [dict(row) for row in cas.execute(
        "SELECT cas_no,COUNT(*) AS n FROM cas_substance WHERE TRIM(cas_no)<>'' "
        "GROUP BY cas_no HAVING COUNT(*)>1")]
    if duplicate_cas:
        errors.append(f"有效 CAS 重复: {duplicate_cas}")
    bad_checksum = [row["cas_no"] for row in valid_rows
                    if not _cas_checksum_ok(str(row["cas_no"]).strip())]
    if bad_checksum:
        errors.append(f"CAS 校验位错误: {bad_checksum}")

    cas_without_models = [row["cas_no"] for row in cas.execute(
        "SELECT s.cas_no FROM cas_substance s LEFT JOIN cas_model_usage u ON u.cas_id=s.cas_id "
        "WHERE TRIM(s.cas_no)<>'' AND u.cas_id IS NULL")]
    if cas_without_models:
        errors.append(f"已有 CAS 没有关联型号: {cas_without_models}")
    raw_cas_orphans = cas.execute(
        "SELECT COUNT(*) FROM cas_model_usage u LEFT JOIN cas_substance s "
        "ON TRIM(s.cas_no)=TRIM(u.raw_cas) WHERE TRIM(u.raw_cas)<>'' AND s.cas_id IS NULL"
    ).fetchone()[0]
    if raw_cas_orphans:
        errors.append(f"原始 CAS 无法回到主表: {raw_cas_orphans} 条")
    raw_cas_multiple_ids = cas.execute(
        "SELECT COUNT(*) FROM (SELECT raw_cas FROM cas_model_usage WHERE TRIM(raw_cas)<>'' "
        "GROUP BY raw_cas HAVING COUNT(DISTINCT cas_id)>1)").fetchone()[0]
    if raw_cas_multiple_ids:
        errors.append(f"同一原始 CAS 对应多个 cas_id: {raw_cas_multiple_ids} 个")
    fk_orphans = {
        "alias": cas.execute(
            "SELECT COUNT(*) FROM cas_alias a LEFT JOIN cas_substance s ON s.cas_id=a.cas_id "
            "WHERE s.cas_id IS NULL").fetchone()[0],
        "usage": cas.execute(
            "SELECT COUNT(*) FROM cas_model_usage u LEFT JOIN cas_substance s ON s.cas_id=u.cas_id "
            "WHERE s.cas_id IS NULL").fetchone()[0],
    }
    if any(fk_orphans.values()):
        errors.append(f"CAS 外键孤儿记录: {fk_orphans}")

    source_count_mismatch = [dict(row) for row in cas.execute(
        "SELECT s.cas_id,s.cas_no,s.standard_name,s.source_count,"
        "COUNT(DISTINCT u.source_file) AS actual_source_count "
        "FROM cas_substance s LEFT JOIN cas_model_usage u ON u.cas_id=s.cas_id "
        "GROUP BY s.cas_id HAVING s.source_count<>COUNT(DISTINCT u.source_file)")]
    if source_count_mismatch:
        errors.append(f"source_count 不是来源文件去重数: {source_count_mismatch}")
    model_count_mismatch = [dict(row) for row in cas.execute(
        "SELECT s.cas_id,s.cas_no,s.standard_name,s.model_count,"
        "COUNT(DISTINCT u.model) AS actual_model_count "
        "FROM cas_substance s LEFT JOIN cas_model_usage u ON u.cas_id=s.cas_id "
        "GROUP BY s.cas_id HAVING s.model_count<>COUNT(DISTINCT u.model)")]
    if model_count_mismatch:
        errors.append(f"model_count 与实际型号去重数不一致: {model_count_mismatch}")

    model_result = {}
    if model_db and model_db.exists():
        model = _ro(model_db)
        all_models = {row[0] for row in model.execute("SELECT model FROM msds_model")}
        used_models = {row[0] for row in cas.execute("SELECT DISTINCT model FROM cas_model_usage")}
        missing_models = sorted(all_models - used_models)
        model_result = {
            "total_models": len(all_models),
            "cas_usage_models": len(used_models),
            "models_without_component_usage": missing_models,
        }
        if missing_models:
            warnings.append(f"总型号库存在无 Section 3 成分关联型号: {missing_models}")
        model.close()

    mapping_result = {}
    if mapping_db:
        if not mapping_db.exists():
            errors.append(f"独立 CAS 字段映射库不存在: {mapping_db}")
        else:
            mapping = _ro(mapping_db)
            mapping_tables = {row["name"] for row in mapping.execute(
                "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")}
            required_mapping_tables = {
                "cas_identity", "cas_field_mapping", "cas_mapping_observation",
                "cas_mapping_conflict",
            }
            missing_mapping_tables = sorted(required_mapping_tables - mapping_tables)
            if missing_mapping_tables:
                errors.append(f"CAS 字段映射库缺少表: {missing_mapping_tables}")
            active_identity_cas_dup = mapping.execute(
                "SELECT COUNT(*) FROM (SELECT cas_no FROM cas_identity WHERE status='active' "
                "AND TRIM(cas_no)<>'' GROUP BY cas_no HAVING COUNT(*)>1)").fetchone()[0]
            if active_identity_cas_dup:
                errors.append(f"映射库有效 CAS 重复: {active_identity_cas_dup} 个")
            observation_count = mapping.execute(
                "SELECT COUNT(*) FROM cas_mapping_observation").fetchone()[0]
            usage_count = cas.execute("SELECT COUNT(*) FROM cas_model_usage").fetchone()[0]
            if observation_count != usage_count:
                errors.append(f"映射观察记录 {observation_count} 与 CAS 使用关系 {usage_count} 不一致")
            mapping_result = {
                "identity_count": mapping.execute("SELECT COUNT(*) FROM cas_identity").fetchone()[0],
                "active_identity_count": mapping.execute(
                    "SELECT COUNT(*) FROM cas_identity WHERE status='active'").fetchone()[0],
                "field_mapping_count": mapping.execute(
                    "SELECT COUNT(*) FROM cas_field_mapping").fetchone()[0],
                "observation_count": observation_count,
                "conflict_count": mapping.execute(
                    "SELECT COUNT(*) FROM cas_mapping_conflict").fetchone()[0],
            }
            mapping.close()

    result = {
        "passed": not errors,
        "errors": errors,
        "warnings": warnings,
        "cas": {
            "substances": len(substances),
            "valid_cas": len(valid_rows),
            "no_cas": len(substances) - len(valid_rows),
            "aliases": cas.execute("SELECT COUNT(*) FROM cas_alias").fetchone()[0],
            "usages": cas.execute("SELECT COUNT(*) FROM cas_model_usage").fetchone()[0],
            "valid_cas_without_models": len(cas_without_models),
            "raw_cas_orphans": raw_cas_orphans,
            "source_count_mismatch": len(source_count_mismatch),
        },
        "models": model_result,
        "mapping": mapping_result,
    }
    cas.close()
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description="验证 CAS 身份库和独立映射库")
    parser.add_argument("cas_db", type=Path)
    parser.add_argument("--mapping-db", type=Path)
    parser.add_argument("--model-db", type=Path)
    args = parser.parse_args()
    result = validate(args.cas_db, args.mapping_db, args.model_db)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
