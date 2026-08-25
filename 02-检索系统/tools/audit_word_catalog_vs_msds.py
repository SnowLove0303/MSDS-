# -*- coding: utf-8 -*-
"""审计指定 WORD 版本目录的型号清单与正式中文 MSDS 库差异。"""
from __future__ import annotations

import argparse
import json
import re
import sqlite3
from pathlib import Path

ROOT = Path(r"F:\正式项目与模块化内容\冠志\MSDS")
CATALOG_ROOT = Path(r"F:\冠志工作空间\产品\TDS MSDS\TDS MSDS\产品 TDS MSDS -- WORD版本")
DB = ROOT / "03-数据库" / "正式库" / "Data Base" / "msds_standard.db"
CODE_RE = re.compile(r"(?<![A-Za-z0-9])([A-Z]{1,6}(?:[-_ ]?\d{2,}[A-Z0-9]*|\d{2,}[A-Z0-9]*))(?![A-Za-z0-9])", re.I)


def normalize(raw: str) -> str:
    value = raw.upper().replace("_", "-").replace(" ", "-")
    return re.sub(r"(?:TDS|MSDS|DS)$", "", value)


def audit() -> dict:
    with sqlite3.connect(DB) as conn:
        known = {row[0].upper() for row in conn.execute("SELECT model FROM msds_model")}
    found: dict[str, dict] = {}
    for path in CATALOG_ROOT.rglob("*"):
        if not path.is_file() or path.name.startswith("~$") or path.suffix.lower() not in (".doc", ".docx", ".pdf"):
            continue
        name = path.name.lower()
        if not any(x in name for x in ("msds", "tds")):
            continue
        for match in CODE_RE.finditer(path.stem.upper()):
            model = normalize(match.group(1))
            item = found.setdefault(model, {"model": model, "files": [], "cn_msds": [], "cn_tds": [], "en_msds": []})
            item["files"].append(str(path))
            is_cn = any(x in name for x in ("cn", "中文", "中文版"))
            if "msds" in name and is_cn:
                item["cn_msds"].append(str(path))
            elif "tds" in name and is_cn:
                item["cn_tds"].append(str(path))
            elif "msds" in name and any(x in name for x in ("en", "英文")):
                item["en_msds"].append(str(path))
    for item in found.values():
        model = item["model"]
        item["in_msds_db"] = model in known
        if item["in_msds_db"]:
            item["status"] = "已入中文MSDS库"
        elif item["cn_msds"]:
            item["status"] = "中文MSDS待入库"
        elif item["cn_tds"]:
            item["status"] = "仅中文TDS，不能直接写入MSDS库"
        elif item["en_msds"]:
            item["status"] = "仅英文MSDS"
        else:
            item["status"] = "待人工确认"
    return {
        "catalog_root": str(CATALOG_ROOT),
        "db": str(DB),
        "db_models": len(known),
        "catalog_candidates": len(found),
        "cn_msds_candidates": sum(bool(x["cn_msds"]) for x in found.values()),
        "cn_msds_not_in_db": [x for x in found.values() if x["cn_msds"] and not x["in_msds_db"]],
        "not_in_db": [x for x in found.values() if not x["in_msds_db"]],
        "items": sorted(found.values(), key=lambda x: x["model"]),
    }


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--out", default=str(ROOT / "_codex_work" / "word_catalog_vs_msds_audit.json"))
    args = parser.parse_args()
    report = audit()
    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    Path(args.out).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({k: report[k] for k in ("db_models", "catalog_candidates", "cn_msds_candidates", "cn_msds_not_in_db")}, ensure_ascii=False, indent=2))
