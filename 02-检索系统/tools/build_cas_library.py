# -*- coding: utf-8 -*-
"""从正式中文总型号库生成独立 CAS/成分库。

CAS 库与型号库分离：
  cas_substance   一种规范成分/物质一条主记录
  cas_alias       原文名称、异写、缩写与规范名的别名关系
  cas_model_usage 型号-成分-浓度的来源关系
  cas_note        飞书规范要求保留的特殊说明

当前版本只使用正式中文总型号库的 Section 3 事实数据；
未出现于当前203型号源文件的旧版知识库条目不会被虚构写入。
"""
from __future__ import annotations

import json
import re
import sqlite3
import sys
from collections import Counter, defaultdict
from pathlib import Path

from build_cas_mapping_library import MappingResolver

CAS_RE = re.compile(r"(?<!\d)(\d{2,7}-\d{2}-\d)(?!\d)")
CONC_RE = re.compile(r"含量\s*[:：]\s*(.*)$", re.S)


def init_db(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        PRAGMA foreign_keys=ON;
        DROP TABLE IF EXISTS cas_note;
        DROP TABLE IF EXISTS cas_model_usage;
        DROP TABLE IF EXISTS cas_alias;
        DROP TABLE IF EXISTS cas_substance;
        CREATE TABLE cas_substance (
            cas_id INTEGER PRIMARY KEY,
            cas_no TEXT NOT NULL DEFAULT '',
            standard_name TEXT NOT NULL,
            category TEXT NOT NULL,
            registry_status TEXT NOT NULL,
            source_count INTEGER NOT NULL DEFAULT 0,
            model_count INTEGER NOT NULL DEFAULT 0,
            remarks TEXT NOT NULL DEFAULT '',
            UNIQUE(cas_no, standard_name)
        );
        CREATE TABLE cas_alias (
            alias_id INTEGER PRIMARY KEY,
            cas_id INTEGER NOT NULL REFERENCES cas_substance(cas_id),
            alias TEXT NOT NULL,
            alias_type TEXT NOT NULL DEFAULT '原文表述',
            occurrence_count INTEGER NOT NULL DEFAULT 0,
            UNIQUE(cas_id, alias)
        );
        CREATE TABLE cas_model_usage (
            usage_id INTEGER PRIMARY KEY,
            cas_id INTEGER NOT NULL REFERENCES cas_substance(cas_id),
            model TEXT NOT NULL,
            raw_name TEXT NOT NULL,
            raw_cas TEXT NOT NULL DEFAULT '',
            concentration TEXT NOT NULL DEFAULT '',
            source_file TEXT NOT NULL DEFAULT '',
            UNIQUE(cas_id, model, raw_name, concentration)
        );
        CREATE TABLE cas_note (
            note_id INTEGER PRIMARY KEY,
            cas_no TEXT NOT NULL,
            note_type TEXT NOT NULL,
            note TEXT NOT NULL,
            source TEXT NOT NULL DEFAULT '飞书知识库'
        );
        CREATE INDEX idx_cas_substance_no ON cas_substance(cas_no);
        CREATE UNIQUE INDEX ux_cas_substance_no_nonblank
            ON cas_substance(cas_no) WHERE TRIM(cas_no) <> '';
        CREATE INDEX idx_cas_alias_alias ON cas_alias(alias);
        CREATE INDEX idx_cas_usage_model ON cas_model_usage(model);
        """
    )


def is_polymer_or_confidential(name: str) -> bool:
    return any(word in name for word in ("聚", "树脂", "聚合物", "成分", "机密", "低聚物", "乳液"))


def main() -> int:
    if len(sys.argv) < 3:
        print("用法: python build_cas_library.py <型号库.db> <CAS库.db> [CAS字段映射库.db]")
        return 2
    model_db = Path(sys.argv[1])
    cas_db = Path(sys.argv[2])
    mapping_db = Path(sys.argv[3]) if len(sys.argv) >= 4 else Path(
        r"F:\正式项目与模块化内容\冠志\MSDS\03-数据库\正式库\Data Base\cas_field_mapping.db")
    if not model_db.exists():
        raise FileNotFoundError(model_db)
    if not mapping_db.exists():
        raise FileNotFoundError(mapping_db)

    source = sqlite3.connect(f"file:{model_db.as_posix()}?mode=ro", uri=True)
    source.row_factory = sqlite3.Row
    resolver = MappingResolver(mapping_db)
    rows = list(source.execute(
        """SELECT m.model,m.source_file,f.label,f.value
           FROM msds_field f JOIN msds_model m ON m.model_id=f.model_id
           WHERE f.section=3 AND f.kind='component'
           ORDER BY m.model,f.id"""))

    parsed = []
    for row in rows:
        raw_name = str(row["label"] or "").strip()
        raw_value = str(row["value"] or "").strip()
        cas_numbers = CAS_RE.findall(raw_value) or CAS_RE.findall(raw_name)
        cas_no = cas_numbers[0] if cas_numbers else ""
        m = CONC_RE.search(raw_value)
        concentration = m.group(1).strip() if m else raw_value
        parsed.append({"model": row["model"], "source_file": row["source_file"],
                       "raw_name": raw_name, "raw_value": raw_value,
                       "cas_no": cas_no, "concentration": concentration})

    conn = sqlite3.connect(cas_db)
    init_db(conn)
    by_key: dict[str, list[dict]] = defaultdict(list)
    identity_meta: dict[str, dict[str, str]] = {}
    for item in parsed:
        resolved = resolver.resolve(item["cas_no"], item["raw_name"])
        if resolved:
            key = str(resolved["identity_key"])
            standard = str(resolved["standard_name"] or item["raw_name"])
            mapping_type = str(resolved.get("mapping_type") or (
                "规范名" if item["raw_name"] == standard else "原文表述/异写"))
        else:
            key = (f"CAS:{item['cas_no']}" if item["cas_no"] else
                   f"NO_CAS:{item['raw_name']}")
            standard = item["raw_name"]
            mapping_type = "规范名" if item["raw_name"] == standard else "待确认"
        identity_meta.setdefault(key, {
            "cas_no": item["cas_no"],
            "standard": standard,
            "mapping_type": mapping_type,
        })
        by_key[key].append(item)

    cas_id_by_identity: dict[str, int] = {}
    for key, items in sorted(by_key.items()):
        meta = identity_meta[key]
        cas_no = meta["cas_no"]
        standard = meta["standard"]
        category = "明确CAS成分" if cas_no else (
            "高分子主体/保密成分" if is_polymer_or_confidential(standard) else "其他无CAS成分")
        status = "已确认CAS" if cas_no else (
            "商业机密/聚合物未设CAS" if is_polymer_or_confidential(standard) else "待确认")
        cur = conn.execute(
            "INSERT INTO cas_substance(cas_no,standard_name,category,registry_status,source_count,model_count) VALUES(?,?,?,?,?,?)",
            (cas_no, standard, category, status, len({x['source_file'] for x in items}),
             len({x['model'] for x in items})))
        cas_id = cur.lastrowid
        cas_id_by_identity[key] = cas_id
        aliases: dict[str, int] = defaultdict(int)
        for item in items:
            aliases[item["raw_name"]] += 1
        for raw_name, occurrence_count in aliases.items():
            alias_type = ("规范名" if raw_name == standard else
                          meta.get("mapping_type") or "原文表述/异写")
            conn.execute(
                "INSERT INTO cas_alias(cas_id,alias,alias_type,occurrence_count) VALUES(?,?,?,?)",
                (cas_id, raw_name, alias_type, occurrence_count))
        for item in items:
            conn.execute(
                "INSERT INTO cas_model_usage(cas_id,model,raw_name,raw_cas,concentration,source_file) VALUES(?,?,?,?,?,?)",
                (cas_id, item["model"], item["raw_name"], item["cas_no"], item["concentration"], item["source_file"]))

    special_notes = {
        "108-01-0": "中和剂；飞书知识库记载部分原文为已键合为盐，质量浓度小于2.0%，具体以型号原始Section 3说明为准。",
        "121-44-8": "聚氨酯阴离子自乳化体系中作为中和剂；具体成盐说明以型号原始Section 3为准。",
    }
    for cas_no, note in special_notes.items():
        if conn.execute("SELECT 1 FROM cas_substance WHERE cas_no=?", (cas_no,)).fetchone():
            conn.execute("INSERT INTO cas_note(cas_no,note_type,note) VALUES(?,?,?)", (cas_no, "特殊中和剂/成盐说明", note))

    # 统计值必须是实际来源文件数，而不是成分使用关系行数。
    conn.execute("""UPDATE cas_substance SET
        source_count=(SELECT COUNT(DISTINCT source_file) FROM cas_model_usage u WHERE u.cas_id=cas_substance.cas_id),
        model_count=(SELECT COUNT(DISTINCT model) FROM cas_model_usage u WHERE u.cas_id=cas_substance.cas_id)""")

    duplicate_cas = conn.execute(
        "SELECT cas_no FROM cas_substance WHERE TRIM(cas_no)<>'' "
        "GROUP BY cas_no HAVING COUNT(*)>1"
    ).fetchall()
    without_model = conn.execute(
        "SELECT s.cas_no FROM cas_substance s LEFT JOIN cas_model_usage u ON u.cas_id=s.cas_id "
        "WHERE TRIM(s.cas_no)<>'' AND u.cas_id IS NULL"
    ).fetchall()
    raw_cas_orphans = conn.execute(
        "SELECT COUNT(*) FROM cas_model_usage u LEFT JOIN cas_substance s "
        "ON TRIM(s.cas_no)=TRIM(u.raw_cas) WHERE TRIM(u.raw_cas)<>'' AND s.cas_id IS NULL"
    ).fetchone()[0]
    source_count_mismatch = conn.execute(
        "SELECT COUNT(*) FROM (SELECT s.cas_id FROM cas_substance s "
        "LEFT JOIN cas_model_usage u ON u.cas_id=s.cas_id GROUP BY s.cas_id "
        "HAVING s.source_count<>COUNT(DISTINCT u.source_file))"
    ).fetchone()[0]
    if duplicate_cas or without_model or raw_cas_orphans or source_count_mismatch:
        raise RuntimeError(json.dumps({
            "duplicate_cas": [row[0] for row in duplicate_cas],
            "cas_without_model": [row[0] for row in without_model],
            "raw_cas_orphans": raw_cas_orphans,
            "source_count_mismatch": source_count_mismatch,
        }, ensure_ascii=False))
    conn.commit()
    counts = {
        "component_rows": len(parsed),
        "substances": conn.execute("SELECT COUNT(*) FROM cas_substance").fetchone()[0],
        "cas_confirmed": conn.execute("SELECT COUNT(*) FROM cas_substance WHERE cas_no<>''").fetchone()[0],
        "no_cas": conn.execute("SELECT COUNT(*) FROM cas_substance WHERE cas_no='' ").fetchone()[0],
        "aliases": conn.execute("SELECT COUNT(*) FROM cas_alias").fetchone()[0],
        "usages": conn.execute("SELECT COUNT(*) FROM cas_model_usage").fetchone()[0],
    }
    print(json.dumps({**counts, "mapping_db": str(mapping_db)}, ensure_ascii=False, indent=2))
    resolver.close()
    source.close()
    conn.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
