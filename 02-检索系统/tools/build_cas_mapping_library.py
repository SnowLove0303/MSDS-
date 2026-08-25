# -*- coding: utf-8 -*-
"""构建独立 CAS 字段映射库。

CAS 身份库只保存物质身份、别名、型号使用关系和来源；本库保存
CAS/无 CAS 成分的规范名、原始名称、异写、观察记录与名称冲突。

该脚本只从现有正式 CAS 身份库和历史构建脚本的规范名映射读取事实，
默认要求目标库不存在，避免误覆盖正式映射库。需要重建时应先备份，
再显式传入 --replace。
"""
from __future__ import annotations

import argparse
import ast
import json
import sqlite3
from collections import defaultdict
from datetime import datetime
from pathlib import Path


DEFAULT_ROOT = Path(r"F:\正式项目与模块化内容\冠志\MSDS")
DEFAULT_CAS_DB = DEFAULT_ROOT / r"03-数据库\正式库\Data Base\cas_library.db"
DEFAULT_MAPPING_DB = DEFAULT_ROOT / r"03-数据库\正式库\Data Base\cas_field_mapping.db"


SCHEMA = """
PRAGMA foreign_keys=ON;
CREATE TABLE cas_identity (
    identity_key TEXT PRIMARY KEY,
    cas_no TEXT NOT NULL DEFAULT '',
    cas_id_ref INTEGER NOT NULL DEFAULT 0,
    standard_name TEXT NOT NULL,
    identity_type TEXT NOT NULL,
    category TEXT NOT NULL DEFAULT '',
    registry_status TEXT NOT NULL DEFAULT '',
    status TEXT NOT NULL DEFAULT 'active',
    source TEXT NOT NULL DEFAULT '',
    remarks TEXT NOT NULL DEFAULT ''
);
CREATE UNIQUE INDEX ux_cas_identity_cas_nonblank
    ON cas_identity(cas_no) WHERE TRIM(cas_no) <> '';
CREATE TABLE cas_field_mapping (
    mapping_id INTEGER PRIMARY KEY,
    identity_key TEXT NOT NULL REFERENCES cas_identity(identity_key),
    cas_no TEXT NOT NULL DEFAULT '',
    raw_name TEXT NOT NULL,
    standard_name TEXT NOT NULL,
    mapping_type TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'active',
    occurrence_count INTEGER NOT NULL DEFAULT 0,
    source TEXT NOT NULL DEFAULT '',
    first_seen_model TEXT NOT NULL DEFAULT '',
    remarks TEXT NOT NULL DEFAULT '',
    UNIQUE(identity_key, raw_name)
);
CREATE INDEX idx_cas_mapping_raw_name ON cas_field_mapping(raw_name);
CREATE INDEX idx_cas_mapping_cas_no ON cas_field_mapping(cas_no);
CREATE TABLE cas_mapping_observation (
    observation_id INTEGER PRIMARY KEY,
    mapping_id INTEGER NOT NULL REFERENCES cas_field_mapping(mapping_id),
    model TEXT NOT NULL,
    raw_cas TEXT NOT NULL DEFAULT '',
    concentration TEXT NOT NULL DEFAULT '',
    source_file TEXT NOT NULL DEFAULT '',
    UNIQUE(mapping_id, model, raw_cas, concentration, source_file)
);
CREATE INDEX idx_cas_mapping_observation_model ON cas_mapping_observation(model);
CREATE TABLE cas_mapping_conflict (
    conflict_id INTEGER PRIMARY KEY,
    raw_name TEXT NOT NULL UNIQUE,
    candidate_count INTEGER NOT NULL,
    candidate_identity_keys TEXT NOT NULL,
    candidate_cas_nos TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'ambiguous_without_cas',
    reason TEXT NOT NULL DEFAULT '',
    source TEXT NOT NULL DEFAULT ''
);
CREATE TABLE cas_mapping_meta (
    meta_key TEXT PRIMARY KEY,
    meta_value TEXT NOT NULL DEFAULT ''
);
"""


def _load_legacy_mapping(path: Path) -> dict[str, str]:
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    for node in tree.body:
        if not isinstance(node, ast.Assign):
            continue
        if any(isinstance(target, ast.Name) and target.id == "CAS_STANDARD_NAMES"
               for target in node.targets):
            value = ast.literal_eval(node.value)
            if not isinstance(value, dict):
                raise ValueError("CAS_STANDARD_NAMES 不是字典")
            return {str(k): str(v) for k, v in value.items()}
    raise ValueError(f"未找到 CAS_STANDARD_NAMES: {path}")


def _identity_key(cas_no: str, standard_name: str, cas_id: int) -> str:
    if cas_no:
        return f"CAS:{cas_no}"
    return f"NO_CAS:{standard_name}"


def _insert_or_update_mapping(conn: sqlite3.Connection, *, identity_key: str,
                              cas_no: str, raw_name: str, standard_name: str,
                              mapping_type: str, occurrence_count: int,
                              source: str, first_seen_model: str = "",
                              status: str = "active", remarks: str = "") -> int:
    row = conn.execute(
        "SELECT mapping_id, occurrence_count, first_seen_model, status, remarks "
        "FROM cas_field_mapping WHERE identity_key=? AND raw_name=?",
        (identity_key, raw_name),
    ).fetchone()
    if row:
        first = row[2] or first_seen_model
        current_status = row[3] or status
        current_remarks = row[4] or remarks
        conn.execute(
            "UPDATE cas_field_mapping SET standard_name=?, mapping_type=?, "
            "status=?, occurrence_count=?, source=?, first_seen_model=?, remarks=? "
            "WHERE mapping_id=?",
            (standard_name, mapping_type, current_status,
             max(int(row[1] or 0), int(occurrence_count)), source, first, current_remarks, row[0]),
        )
        return int(row[0])
    cur = conn.execute(
        "INSERT INTO cas_field_mapping(" 
        "identity_key,cas_no,raw_name,standard_name,mapping_type,status,"
        "occurrence_count,source,first_seen_model,remarks) VALUES(?,?,?,?,?,?,?,?,?,?)",
        (identity_key, cas_no, raw_name, standard_name, mapping_type, status,
         occurrence_count, source, first_seen_model, remarks),
    )
    return int(cur.lastrowid)


class MappingResolver:
    """建库时读取独立映射库，不把名称映射写回身份主表。"""

    def __init__(self, mapping_db: Path):
        self.connection = sqlite3.connect(f"file:{mapping_db.as_posix()}?mode=ro", uri=True)
        self.connection.row_factory = sqlite3.Row

    def resolve(self, cas_no: str, raw_name: str) -> dict[str, str] | None:
        cas_no = str(cas_no or "").strip()
        raw_name = str(raw_name or "").strip()
        if cas_no:
            row = self.connection.execute(
                "SELECT i.identity_key,i.cas_no,i.standard_name,m.mapping_type "
                "FROM cas_identity i LEFT JOIN cas_field_mapping m "
                "ON m.identity_key=i.identity_key AND m.raw_name=i.standard_name "
                "WHERE i.cas_no=? AND i.status='active' ORDER BY m.mapping_id LIMIT 1",
                (cas_no,),
            ).fetchone()
            if not row:
                return None
            return dict(row)

        rows = self.connection.execute(
            "SELECT DISTINCT identity_key,cas_no,standard_name,mapping_type "
            "FROM cas_field_mapping WHERE cas_no='' AND raw_name=? AND status='active' "
            "ORDER BY mapping_id",
            (raw_name,),
        ).fetchall()
        if len(rows) != 1:
            return None
        return dict(rows[0])

    def close(self) -> None:
        self.connection.close()


def build(cas_db: Path, mapping_db: Path, legacy_builder: Path | None, replace: bool) -> dict:
    if not cas_db.exists():
        raise FileNotFoundError(cas_db)
    if legacy_builder and not legacy_builder.exists():
        raise FileNotFoundError(legacy_builder)
    if mapping_db.exists() and not replace:
        raise FileExistsError(f"目标映射库已存在，若要重建请显式使用 --replace: {mapping_db}")
    mapping_db.parent.mkdir(parents=True, exist_ok=True)
    if mapping_db.exists():
        mapping_db.unlink()

    source = sqlite3.connect(f"file:{cas_db.as_posix()}?mode=ro", uri=True)
    source.row_factory = sqlite3.Row
    legacy = _load_legacy_mapping(legacy_builder) if legacy_builder else {}
    conn = sqlite3.connect(mapping_db)
    conn.row_factory = sqlite3.Row
    conn.executescript(SCHEMA)

    substances = source.execute(
        "SELECT cas_id,cas_no,standard_name,category,registry_status "
        "FROM cas_substance ORDER BY cas_id"
    ).fetchall()
    usage_by_cas = {
        int(row["cas_id"]): dict(row)
        for row in source.execute(
            "SELECT cas_id,COUNT(*) AS usage_count,MIN(model) AS first_model "
            "FROM cas_model_usage GROUP BY cas_id"
        )
    }
    for row in substances:
        cas_no = str(row["cas_no"] or "").strip()
        name = str(row["standard_name"] or "").strip()
        key = _identity_key(cas_no, name, int(row["cas_id"]))
        usage = usage_by_cas.get(int(row["cas_id"]), {"usage_count": 0, "first_model": ""})
        conn.execute(
            "INSERT INTO cas_identity(identity_key,cas_no,cas_id_ref,standard_name,"
            "identity_type,category,registry_status,status,source) VALUES(?,?,?,?,?,?,?,?,?)",
            (key, cas_no, int(row["cas_id"]), name,
             "CAS" if cas_no else "NO_CAS", row["category"], row["registry_status"],
             "active", "正式 CAS 身份库"),
        )
        _insert_or_update_mapping(
            conn, identity_key=key, cas_no=cas_no, raw_name=name,
            standard_name=name, mapping_type="规范名",
            occurrence_count=int(usage["usage_count"] or 0), source="正式 CAS 身份库",
            first_seen_model=str(usage["first_model"] or ""),
        )

    aliases = source.execute(
        "SELECT a.cas_id,a.alias,a.alias_type,a.occurrence_count,"
        "s.cas_no,s.standard_name FROM cas_alias a JOIN cas_substance s ON s.cas_id=a.cas_id "
        "ORDER BY a.alias_id"
    ).fetchall()
    identity_by_cas_id = {
        int(row["cas_id"]): _identity_key(str(row["cas_no"] or "").strip(),
                                           str(row["standard_name"] or "").strip(),
                                           int(row["cas_id"]))
        for row in substances
    }
    for row in aliases:
        key = identity_by_cas_id[int(row["cas_id"])]
        first = source.execute(
            "SELECT MIN(model) FROM cas_model_usage WHERE cas_id=? AND raw_name=?",
            (row["cas_id"], row["alias"]),
        ).fetchone()[0] or ""
        _insert_or_update_mapping(
            conn, identity_key=key, cas_no=str(row["cas_no"] or "").strip(),
            raw_name=str(row["alias"] or "").strip(),
            standard_name=str(row["standard_name"] or "").strip(),
            mapping_type=str(row["alias_type"] or "原文表述/异写"),
            occurrence_count=int(row["occurrence_count"] or 0), source="正式 CAS 身份库",
            first_seen_model=str(first),
        )

    current_cas = {str(row["cas_no"] or "").strip() for row in substances
                   if str(row["cas_no"] or "").strip()}
    for cas_no, standard_name in sorted(legacy.items()):
        if cas_no in current_cas:
            continue
        key = _identity_key(cas_no, standard_name, 0)
        conn.execute(
            "INSERT INTO cas_identity(identity_key,cas_no,cas_id_ref,standard_name,"
            "identity_type,status,source,remarks) VALUES(?,?,?,?,?,?,?,?)",
            (key, cas_no, 0, standard_name, "CAS", "pending",
             "历史构建脚本 CAS_STANDARD_NAMES",
             "当前正式 CAS 身份库未发现型号使用关系，暂不写入正式身份库"),
        )
        _insert_or_update_mapping(
            conn, identity_key=key, cas_no=cas_no, raw_name=standard_name,
            standard_name=standard_name, mapping_type="规范名", occurrence_count=0,
            source="历史构建脚本 CAS_STANDARD_NAMES", status="pending",
            remarks="无当前型号事实，待后续核验",
        )

    mapping_ids = {
        (row["identity_key"], row["raw_name"]): int(row["mapping_id"])
        for row in conn.execute("SELECT mapping_id,identity_key,raw_name FROM cas_field_mapping")
    }
    for row in source.execute(
        "SELECT u.model,u.raw_name,u.raw_cas,u.concentration,u.source_file,"
        "s.cas_id,s.cas_no,s.standard_name FROM cas_model_usage u "
        "JOIN cas_substance s ON s.cas_id=u.cas_id ORDER BY u.usage_id"
    ):
        key = identity_by_cas_id[int(row["cas_id"])]
        map_id = mapping_ids.get((key, str(row["raw_name"] or "").strip()))
        if map_id is None:
            map_id = _insert_or_update_mapping(
                conn, identity_key=key, cas_no=str(row["cas_no"] or "").strip(),
                raw_name=str(row["raw_name"] or "").strip(),
                standard_name=str(row["standard_name"] or "").strip(),
                mapping_type="原文表述/异写", occurrence_count=0,
                source="正式 CAS 身份库补齐",
            )
            mapping_ids[(key, str(row["raw_name"] or "").strip())] = map_id
        conn.execute(
            "INSERT OR IGNORE INTO cas_mapping_observation(mapping_id,model,raw_cas,"
            "concentration,source_file) VALUES(?,?,?,?,?)",
            (map_id, row["model"], row["raw_cas"] or "", row["concentration"] or "",
             row["source_file"] or ""),
        )

    conflicts = source.execute(
        "SELECT raw_name,COUNT(DISTINCT cas_id) AS candidate_count "
        "FROM cas_model_usage WHERE TRIM(raw_name)<>'' GROUP BY raw_name "
        "HAVING COUNT(DISTINCT cas_id)>1 ORDER BY raw_name"
    ).fetchall()
    for row in conflicts:
        candidates = source.execute(
            "SELECT DISTINCT s.cas_no, s.standard_name, s.cas_id "
            "FROM cas_model_usage u JOIN cas_substance s ON s.cas_id=u.cas_id "
            "WHERE u.raw_name=? ORDER BY s.cas_id", (row["raw_name"],)
        ).fetchall()
        keys = [identity_by_cas_id[int(x["cas_id"])] for x in candidates]
        cas_nos = [str(x["cas_no"] or "") for x in candidates]
        conn.execute(
            "INSERT INTO cas_mapping_conflict(raw_name,candidate_count,"
            "candidate_identity_keys,candidate_cas_nos,reason,source) VALUES(?,?,?,?,?,?)",
            (row["raw_name"], int(row["candidate_count"]), json.dumps(keys, ensure_ascii=False),
             json.dumps(cas_nos, ensure_ascii=False),
             "同一原始名称对应多个身份；必须结合原始 CAS/型号/来源文件解析",
             "cas_model_usage"),
        )

    meta = {
        "schema_version": "1.0.0",
        "built_at": datetime.now().isoformat(timespec="seconds"),
        "cas_source": str(cas_db),
        "legacy_mapping_source": str(legacy_builder or ""),
        "formal_identity_authority": "cas_library.db",
        "legal_fact_database": "separate_future_database",
    }
    conn.executemany(
        "INSERT INTO cas_mapping_meta(meta_key,meta_value) VALUES(?,?)", meta.items()
    )
    conn.commit()
    result = {
        "identity_count": conn.execute("SELECT COUNT(*) FROM cas_identity").fetchone()[0],
        "active_identity_count": conn.execute("SELECT COUNT(*) FROM cas_identity WHERE status='active'").fetchone()[0],
        "mapping_count": conn.execute("SELECT COUNT(*) FROM cas_field_mapping").fetchone()[0],
        "observation_count": conn.execute("SELECT COUNT(*) FROM cas_mapping_observation").fetchone()[0],
        "conflict_count": conn.execute("SELECT COUNT(*) FROM cas_mapping_conflict").fetchone()[0],
        "legacy_mapping_count": len(legacy),
        "mapping_db": str(mapping_db),
    }
    conn.close()
    source.close()
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description="构建独立 CAS 字段映射库")
    parser.add_argument("--cas-db", type=Path, default=DEFAULT_CAS_DB)
    parser.add_argument("--mapping-db", type=Path, default=DEFAULT_MAPPING_DB)
    parser.add_argument("--legacy-builder", type=Path,
                        help="仅一次性迁移历史硬编码映射时传入旧版构建脚本；正式重建默认不读取")
    parser.add_argument("--replace", action="store_true")
    args = parser.parse_args()
    print(json.dumps(build(args.cas_db, args.mapping_db, args.legacy_builder, args.replace),
                     ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
