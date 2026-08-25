# -*- coding: utf-8 -*-
"""按产品 TDS/MSDS WORD 版本目录的父子级结构建立型号大类索引。"""
from __future__ import annotations

import json
import re
import sqlite3
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(r"F:\正式项目与模块化内容\冠志\MSDS")
DB = ROOT / "03-数据库" / "正式库" / "Data Base" / "msds_standard.db"
CATEGORY_ROOT = Path(r"F:\冠志工作空间\产品\TDS MSDS\TDS MSDS\产品 TDS MSDS -- WORD版本")


def _token_pattern(model: str) -> re.Pattern:
    escaped = re.escape(model.strip())
    return re.compile(rf"(?<![A-Za-z0-9]){escaped}(?![A-Za-z0-9])", re.I)


def build() -> dict:
    with sqlite3.connect(DB) as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS model_category (
                model_id INTEGER PRIMARY KEY REFERENCES msds_model(model_id),
                category_parent TEXT NOT NULL DEFAULT '', category_child TEXT NOT NULL DEFAULT '',
                category_path TEXT NOT NULL DEFAULT '', source_root TEXT NOT NULL DEFAULT '',
                matched_paths TEXT NOT NULL DEFAULT '[]', updated_at TEXT DEFAULT (datetime('now', 'localtime'))
            )
        """)
        models = conn.execute("SELECT model_id, model FROM msds_model ORDER BY model").fetchall()
        candidates = [p for p in CATEGORY_ROOT.rglob("*") if p.is_file()]
        relative_paths = [(p, p.relative_to(CATEGORY_ROOT).parts) for p in candidates]
        result = {"root": str(CATEGORY_ROOT), "models": len(models), "matched": 0, "unmatched": []}
        for model_id, model in models:
            pattern = _token_pattern(model)
            matches = []
            for path, parts in relative_paths:
                if pattern.search(path.name) or pattern.search(str(path)):
                    if parts:
                        matches.append((path, parts))
            # 历史文件名存在 PA-3110tds / PA-4902ds 这类无分隔符后缀。
            # 只有严格边界完全没有命中时才启用包含匹配，避免短型号误命中长型号。
            if not matches:
                loose = model.strip().lower()
                matches = [(path, parts) for path, parts in relative_paths
                           if loose in str(path).lower() and parts]
            if not matches:
                parent = "其他来源（未在WORD版本目录匹配）"
                conn.execute("""
                    INSERT INTO model_category(model_id, category_parent, category_child, category_path, source_root, matched_paths, updated_at)
                    VALUES(?,?,?,?,?,?,datetime('now','localtime'))
                    ON CONFLICT(model_id) DO UPDATE SET category_parent=excluded.category_parent,
                      category_child=excluded.category_child, category_path=excluded.category_path,
                      source_root=excluded.source_root, matched_paths=excluded.matched_paths,
                      updated_at=excluded.updated_at
                """, (model_id, parent, model, "", str(CATEGORY_ROOT), "[]"))
                result["unmatched"].append(model)
                continue
            parent_counts = Counter(parts[0] for _, parts in matches if parts)
            parent = sorted(parent_counts, key=lambda x: (-parent_counts[x], x))[0]
            # UI 中的子节点就是型号本身；category_path 保留原始父目录，便于追溯。
            matched = [str(p) for p, parts in matches if parts[0] == parent][:20]
            conn.execute("""
                INSERT INTO model_category(model_id, category_parent, category_child, category_path, source_root, matched_paths, updated_at)
                VALUES(?,?,?,?,?,?,datetime('now','localtime'))
                ON CONFLICT(model_id) DO UPDATE SET
                  category_parent=excluded.category_parent, category_child=excluded.category_child,
                  category_path=excluded.category_path, source_root=excluded.source_root,
                  matched_paths=excluded.matched_paths, updated_at=excluded.updated_at
            """, (model_id, parent, model, parent, str(CATEGORY_ROOT), json.dumps(matched, ensure_ascii=False)))
            result["matched"] += 1
        conn.commit()
    return result


if __name__ == "__main__":
    print(json.dumps(build(), ensure_ascii=False, indent=2))
