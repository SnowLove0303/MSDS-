# -*- coding: utf-8 -*-
"""同步全库 Section 3 成分标准化名称到 msds_field.std_name."""

import sqlite3
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from core.s3_component_std import standardize_component_name
from core.msds_db import open_db

DB_PATH = r"F:\正式项目与模块化内容\冠志\MSDS\03-数据库\正式库\Data Base\msds_standard.db"

def sync_components():
    conn = open_db(DB_PATH)
    cur = conn.cursor()

    cur.execute("SELECT id, label, value FROM msds_field WHERE section = 3 AND kind = 'component'")
    rows = cur.fetchall()

    updated = 0
    sample_diffs = []
    for fid, label, val in rows:
        cas = ""
        for part in (val or "").split(" | "):
            if part.startswith("CAS:"):
                cas = part[4:].strip()
        std_c = standardize_component_name(label, cas)
        if std_c != label and len(sample_diffs) < 10:
            sample_diffs.append((label, std_c, cas))
        cur.execute("UPDATE msds_field SET std_name = ? WHERE id = ?", (std_c, fid))
        updated += 1

    conn.commit()
    print(f"✅ 成功同步 {updated} 条成分记录的标准化名称 (msds_field.std_name)")
    print("\n标准化收敛样本对比 (前 10 例):")
    for raw, std, cas in sample_diffs:
        print(f"  • 原文: 「{raw}」 (CAS: {cas}) → 标准化: 【{std}】")

if __name__ == "__main__":
    sync_components()
