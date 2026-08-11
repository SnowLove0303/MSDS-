# -*- coding: utf-8 -*-
"""测试库构建: 挑 10 份中文冠志 MSDS, 用当前解析逻辑重建三表测试库.

不动原有正式筛选库.db, 输出到独立目录供 GUI 检索显示验证.
"""
from __future__ import annotations
import hashlib, pickle, sys, time
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
for p in (_ROOT, _ROOT / "批量化读取"):
    if str(p) not in sys.path:
        sys.path.insert(0, str(p))

from core.docx_reader import read_msds
from build_screened_db import build_screened_db

# 10 份中文冠志, 覆盖多种模板变体
SAMPLE = [
    "2-苯氧基乙醇 msds_CN 冠志.docx",
    "RA-15000 msds_CN 冠志.docx",
    "PA-4408 msds_CN 冠志.docx",
    "BL-8085 msds_CN 冠志.docx",
    "BL-8124 msds_CN 冠志.docx",
    "EC-1800 msds_CN 冠志.docx",
    "PA-4851 msds_CN 冠志.docx",
    "OS-1310 msds_CN 冠志.docx",
    "HPU-7651 msds_CN 冠志.docx",
    "BEK-500L msds_CN 冠志.docx",
]
SRC = Path(r"F:\正式项目与模块化内容\Word 覆写模块\数据库\MSDS\中文\冠志 guanzhi")
OUT = Path(r"F:\正式项目与模块化内容\Word 覆写模块\数据库\测试库_10份冠志")


def _sha256(p: Path) -> str:
    return hashlib.sha256(p.read_bytes()).hexdigest()


def to_records(r, path: Path) -> dict:
    """ParseResult → 扁平 records (与 _parsed_clean.pkl 同构)."""
    fields = []
    for num, sec in r.sections.items():
        for row in sec.iter_rows():
            if row.kind == "section":
                continue
            fields.append({
                "sec": num,
                "label": row.label,
                "value": row.value,
                "kind": "field" if row.kind == "field" else row.kind,
                "seq": row.seq or "",
            })
    comps = []
    for c in sec.components if False else []:
        pass
    for num, sec in r.sections.items():
        for c in sec.components:
            comps.append({"name": c.name, "cas": c.cas, "conc": c.conc})
    return {
        "file": r.file_name,
        "path": str(path),
        "lang": "zh",
        "vendor": "冠志",
        "model": r.file_name.split(" msds")[0].strip(),
        "sha256": r.sha256,
        "header": r.header,
        "footer": r.footer,
        "fields": fields,
        "components": comps,
        "anomalies": [
            {"s": a.section, "lvl": a.level, "msg": a.message, "detail": a.detail}
            for a in r.anomalies
        ],
    }


def main():
    records = []
    for fn in SAMPLE:
        fp = SRC / fn
        if not fp.exists():
            print(f"⚠ 缺文件: {fn}")
            continue
        r = read_msds(fp)
        rec = to_records(r, fp)
        records.append(rec)
        s = r.summary()
        print(f"  {fn}: {s['sections']}节 {s['fields']}字段 {s['components']}成分 异常{s['anomalies']}")
    OUT.mkdir(parents=True, exist_ok=True)
    pkl = OUT / "_parsed_10zhengzhi.pkl"
    with open(pkl, "wb") as f:
        pickle.dump(records, f)
    stats = build_screened_db(pkl, OUT)
    print(f"\n建库完成: {stats}")
    print(f"库文件: {OUT / '正式筛选库.db'}")


if __name__ == "__main__":
    main()
