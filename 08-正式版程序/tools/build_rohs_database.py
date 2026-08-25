from __future__ import annotations

import csv
import sqlite3
from pathlib import Path


ROOT = Path(r"F:\正式项目与模块化内容\冠志\MSDS")
SOURCE = ROOT / "04-推断引擎" / "正式法律法规库" / "物质限制清单" / "EU-RoHS-材料级限用物质清单-2011-65-EU-2015-863.csv"
TARGET = ROOT / "04-推断引擎" / "正式法律法规库" / "物质限制清单" / "EU-RoHS-2011-65-EU-2015-863-限制物质数据库.db"


def main() -> None:
    rows = list(csv.DictReader(SOURCE.open("r", encoding="utf-8-sig", newline="")))
    required = ["序号", "物质类别", "中文名称", "英文名称", "CAS号", "EC号", "均质材料限值", "限值ppm", "材料筛查说明", "法规依据"]
    if not rows or any(any(key not in row for key in required) for row in rows):
        raise RuntimeError("RoHS CSV 字段不完整或没有数据")
    TARGET.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(TARGET) as conn:
        conn.executescript("""
        DROP TABLE IF EXISTS rohs_restricted_substance;
        DROP TABLE IF EXISTS source_metadata;
        CREATE TABLE rohs_restricted_substance (
            "序号" INTEGER PRIMARY KEY,
            "物质类别" TEXT NOT NULL,
            "中文名称" TEXT NOT NULL,
            "英文名称" TEXT NOT NULL,
            "CAS号" TEXT NOT NULL,
            "EC号" TEXT NOT NULL,
            "均质材料限值" TEXT NOT NULL,
            "限值ppm" INTEGER NOT NULL,
            "材料筛查说明" TEXT NOT NULL,
            "法规依据" TEXT NOT NULL
        );
        CREATE INDEX idx_rohs_cas ON rohs_restricted_substance("CAS号");
        CREATE INDEX idx_rohs_ec ON rohs_restricted_substance("EC号");
        CREATE TABLE source_metadata (
            "键" TEXT PRIMARY KEY,
            "值" TEXT NOT NULL
        );
        """)
        conn.executemany(
            'INSERT INTO rohs_restricted_substance VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
            [tuple(row[key].strip() for key in required) for row in rows],
        )
        conn.executemany(
            'INSERT INTO source_metadata VALUES (?, ?)',
            [("source_csv", str(SOURCE)), ("record_count", str(len(rows))),
             ("register", "EU RoHS 2011/65/EU Annex II + 2015/863"),
             ("scope", "均质材料级初筛；物质组/元素索引不能替代最终检测")],
        )
        conn.commit()
    print(f"created {TARGET} with {len(rows)} rows")


if __name__ == "__main__":
    main()
