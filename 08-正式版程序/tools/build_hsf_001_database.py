"""Build the formal HSF 001 register from the Feishu Wiki Sheet snapshot."""

from pathlib import Path
import json
import sqlite3

ROOT = Path(r"F:\正式项目与模块化内容\冠志\MSDS")
DB_PATH = ROOT / "04-推断引擎" / "正式法律法规库" / "物质限制清单" / "HSF-001-有害物质清单.db"
SOURCE_PATH = ROOT / "04-推断引擎" / "正式法律法规库" / "物质限制清单" / "HSF-001-飞书Wiki-原始清单.json"

SOURCE = json.loads(SOURCE_PATH.read_text(encoding="utf-8"))
SOURCE_META = SOURCE["source"]
ROWS = SOURCE["rows"]
LIMITS_NOTE = SOURCE_META["limits_note"]


def build() -> None:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    if DB_PATH.exists():
        DB_PATH.unlink()
    projects = list(dict.fromkeys(row["limit_project"] for row in ROWS))
    with sqlite3.connect(DB_PATH) as conn:
        conn.executescript(
            """
            CREATE TABLE hsf_001_meta (
                meta_key TEXT PRIMARY KEY,
                meta_value TEXT NOT NULL
            );
            CREATE TABLE hsf_001_substance (
                "序号" INTEGER PRIMARY KEY,
                "源表行号" INTEGER NOT NULL,
                "类别" TEXT NOT NULL,
                "中文名称" TEXT NOT NULL,
                "英文名称" TEXT NOT NULL,
                "CAS号" TEXT NOT NULL,
                "源表CAS号" TEXT NOT NULL,
                "CAS映射状态" TEXT NOT NULL,
                "物质级判定" TEXT NOT NULL,
                "法规依据" TEXT NOT NULL,
                "备注" TEXT NOT NULL
            );
            """
        )
        metadata = [
            ("清单名称", "Inventec《无有害物质（HSF）管理规范》HSF 001"),
            ("版本", "HSF 001"),
            ("来源", SOURCE_META["wiki_url"]),
            ("来源Sheet", SOURCE_META["sheet_id"]),
            ("来源修订号", str(SOURCE_META["revision"])),
            ("来源范围", SOURCE_META["range"]),
            ("条目数", str(len(ROWS))),
            ("限制项目数", str(len(projects))),
            ("来源说明", SOURCE_META["source_note"]),
            ("CAS规范化说明", SOURCE_META["normalization_note"]),
            ("判定范围", LIMITS_NOTE),
        ]
        conn.executemany(
            "INSERT INTO hsf_001_meta(meta_key, meta_value) VALUES (?, ?)",
            metadata,
        )
        values = []
        for index, row in enumerate(ROWS, start=1):
            cas = row["cas"]
            raw_cas = row["cas_raw"]
            if cas:
                mapping_status = "已建立CAS映射"
                decision = "不符合（CAS命中）"
            elif raw_cas in ("", "多个CAS"):
                mapping_status = "源表未提供单一CAS"
                decision = "需人工确认（CAS范围）"
            else:
                mapping_status = "源表CAS格式待核对"
                decision = "需人工确认（CAS格式）"
            values.append(
                (
                    index,
                    row["source_row"],
                    row["limit_project"],
                    row["substance"],
                    "",
                    cas,
                    raw_cas,
                    mapping_status,
                    decision,
                    "HSF 001（飞书 Wiki）",
                    f"飞书 Wiki HSF 001 Sheet 源表第 {row['source_row']} 行；{LIMITS_NOTE}",
                )
            )
        conn.executemany(
            """
            INSERT INTO hsf_001_substance
            ("序号", "源表行号", "类别", "中文名称", "英文名称", "CAS号", "源表CAS号",
             "CAS映射状态", "物质级判定", "法规依据", "备注")
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            values,
        )
        conn.commit()
    print(f"built {DB_PATH} ({len(ROWS)} rows, {len(projects)} projects)")


if __name__ == "__main__":
    build()

