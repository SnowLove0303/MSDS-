# -*- coding: utf-8 -*-
"""库自动发现模块：扫描正式库目录 + 配置附加库，自动排除备份/测试/临时库。"""
from __future__ import annotations

import json
import re
import sqlite3
from pathlib import Path
from typing import Any

_PROJECT_ROOT = Path(__file__).resolve().parents[2]  # F:\正式项目与模块化内容\冠志\MSDS
_APP_DIR = Path(__file__).resolve().parents[1]       # 08-正式版程序
_CONFIG_PATH = _APP_DIR / "db_config.json"

# 业务树类型
BT_MODELS = "models"   # 总型号库业务树
BT_CAS = "cas"         # CAS 库业务树


def _load_config() -> dict[str, Any]:
    if _CONFIG_PATH.exists():
        return json.loads(_CONFIG_PATH.read_text(encoding="utf-8"))
    return {"scan_dirs": [], "exclude_patterns": [], "extra_libraries": [],
            "display_names": {}, "business_trees": {}}


def _path_matches_any(path: Path, patterns: list[str]) -> bool:
    name = path.name
    for p in patterns or []:
        if not p:
            continue
        if p.startswith("*"):
            if name.endswith(p[1:]):
                return True
        elif p.endswith("*"):
            if name.startswith(p[:-1]):
                return True
        elif p == name:
            return True
    return False


def _is_real_sqlite(path: Path) -> bool:
    """校验文件确实是 SQLite 数据库（读文件头 magic）。"""
    try:
        with open(path, "rb") as f:
            return f.read(16) == b"SQLite format 3\x00"
    except OSError:
        return False


def discover_libraries() -> list[dict[str, Any]]:
    """返回强关联库清单：[{key, path, name, tables:[{name, rows}], business_tree}]"""
    cfg = _load_config()
    exclude = cfg.get("exclude_patterns", []) or []
    display = cfg.get("display_names", {}) or {}
    biz = cfg.get("business_trees", {}) or {}

    found: dict[str, Path] = {}
    for rel in cfg.get("scan_dirs", []) or []:
        d = _PROJECT_ROOT / rel
        if not d.exists():
            continue
        for f in sorted(d.iterdir()):
            if f.suffix.lower() != ".db":
                continue
            if _path_matches_any(f, exclude):
                continue
            if f.name in found:
                continue
            found[f.name] = f
    # 配置附加库（绝对路径或相对项目根）
    for extra in cfg.get("extra_libraries", []) or []:
        p = Path(extra)
        if not p.is_absolute():
            p = _PROJECT_ROOT / p
        if p.exists() and p.suffix.lower() == ".db" and p.name not in found:
            found[p.name] = p

    libs = []
    for name, path in found.items():
        if not _is_real_sqlite(path):
            continue
        tables = _list_tables(path)
        libs.append({
            "key": name,
            "path": str(path),
            "name": display.get(name, name),
            "filename": name,
            "tables": tables,
            "business_tree": biz.get(name),
        })
    # 按业务树优先排序：总型号库、CAS 库在前，其余按名称
    libs.sort(key=lambda x: (0 if x["business_tree"] == BT_MODELS else
                             1 if x["business_tree"] == BT_CAS else 2, x["name"]))
    return libs


def _list_tables(path: Path) -> list[dict[str, Any]]:
    try:
        conn = sqlite3.connect(f"file:{path}?mode=ro", uri=True)
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' "
            "ORDER BY name").fetchall()
        result = []
        for r in rows:
            tname = r["name"]
            try:
                cnt = conn.execute(f'SELECT COUNT(*) FROM "{tname}"').fetchone()[0]
            except sqlite3.Error:
                cnt = None
            result.append({"name": tname, "rows": cnt})
        conn.close()
        return result
    except sqlite3.Error:
        return []


def open_readonly(path: str) -> sqlite3.Connection:
    conn = sqlite3.connect(f"file:{path}?mode=ro", uri=True)
    conn.row_factory = sqlite3.Row
    return conn
