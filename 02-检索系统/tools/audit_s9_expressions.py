# -*- coding: utf-8 -*-
"""Section 9 原文表述检索与库表记录工具.

扫描数据库及原始 docx 文档，提取 Section 9 的所有原文标签表述，
与标准字段进行映射比对，并在 SQLite 中创建 `s9_label_mapping` 表
和 `s9_raw_expression` 表进行持久化记录。
"""
import sqlite3
import json
import sys
from pathlib import Path
from collections import defaultdict

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from core.docx_reader import read_msds
from core.schema import standard_fields, standard_name
from core.msds_db import _db_std_name, _listed_field_names, _model_of, validate_structure

def run_audit(db_path: str, docx_dir: str):
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()

    # 创建 S9 表述统计表
    cur.execute("""
    CREATE TABLE IF NOT EXISTS s9_label_mapping (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        std_name TEXT NOT NULL,             -- 标准字段名 (如: 闪点, pH值, 密度)
        std_seq TEXT DEFAULT '',            -- 标准骨架序号 (如: 9.6, 9.3, 9.12)
        raw_label TEXT NOT NULL,            -- 原文不同表述 (如: 闪点（℃）, PH值, 相对密度)
        occurrences_count INTEGER DEFAULT 0,-- 出现次数 (命中型号数)
        models_sample TEXT DEFAULT '',      -- 命中型号清单 (逗号分隔)
        sample_values TEXT DEFAULT '',      -- 示例值
        created_at TEXT DEFAULT (datetime('now', 'localtime')),
        UNIQUE(std_name, raw_label)
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS s9_raw_expression (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        model TEXT NOT NULL,                -- 产品型号
        source_file TEXT NOT NULL,          -- 来源文件
        raw_seq TEXT DEFAULT '',            -- 原文序号 (如 9.6, 9.5)
        raw_label TEXT NOT NULL,            -- 原文标签
        std_name TEXT NOT NULL,             -- 归类后的标准字段名
        is_standard INTEGER DEFAULT 1,      -- 是否成功归类进标准骨架
        raw_value TEXT DEFAULT '',          -- 字段原文内容
        created_at TEXT DEFAULT (datetime('now', 'localtime'))
    );
    """)

    # 清空旧统计
    cur.execute("DELETE FROM s9_label_mapping")
    cur.execute("DELETE FROM s9_raw_expression")

    folder = Path(docx_dir)
    files = sorted(f for f in folder.glob("*.docx") if not f.name.startswith("~$"))
    tpl = Path(r"F:\正式项目与模块化内容\冠志\MSDS\03-数据库\正式库\推导方案\PEA-4139 MSDS_CN 冠志 模板.docx")
    if tpl.exists() and tpl not in files:
        files.insert(0, tpl)

    std_fields_9 = {f.name: f for f in standard_fields(9)}
    std_names_s9 = _listed_field_names(9)

    # 骨架标准序号字典 (9.1 ~ 9.23)
    from core.msds_db import _SKELETON
    std_seq_dict = {}
    for item in _SKELETON.get(9, []):
        kind, seq, label = item[0], item[1], item[2]
        std_seq_dict[label] = seq

    mapping_stats = defaultdict(lambda: defaultdict(list))
    raw_rows = []

    for p in files:
        r = read_msds(str(p))
        model = _model_of(r, p.name)
        sec9 = r.sections.get(9)
        if not sec9:
            continue
        cur_major = {}
        for row in sec9.iter_rows():
            if row.kind == "field" and row.label:
                std = _db_std_name(9, row.label, cur_major)
                is_std = std in std_names_s9
                final_std = std if is_std else "(未归类/非标准字段)"
                
                raw_rows.append((
                    model, p.name, row.seq, row.label, final_std,
                    1 if is_std else 0, row.value
                ))
                if is_std:
                    mapping_stats[std][row.label].append((model, row.seq, row.value))

    # 插入明细
    cur.executemany("""
    INSERT INTO s9_raw_expression (model, source_file, raw_seq, raw_label, std_name, is_standard, raw_value)
    VALUES (?, ?, ?, ?, ?, ?, ?)
    """, raw_rows)

    # 插入汇总映射
    mapping_rows = []
    for std_name, labels in sorted(mapping_stats.items()):
        seq = std_seq_dict.get(std_name, "")
        for raw_label, items in sorted(labels.items(), key=lambda x: len(x[1]), reverse=True):
            models = sorted(list({x[0] for x in items}))
            models_str = ", ".join(models)
            # 取 2-3 个非空示例值
            vals = [x[2].strip().replace("\n", " ") for x in items if x[2].strip()]
            sample_val = " | ".join(vals[:2]) if vals else "无数据"
            mapping_rows.append((
                std_name, seq, raw_label, len(models), models_str, sample_val[:200]
            ))

    cur.executemany("""
    INSERT INTO s9_label_mapping (std_name, std_seq, raw_label, occurrences_count, models_sample, sample_values)
    VALUES (?, ?, ?, ?, ?, ?)
    """, mapping_rows)

    conn.commit()
    print(f"✅ Section 9 原文表述检索与库表建立完成！")
    print(f"   - 录入原文表述明细: {len(raw_rows)} 行 (表: s9_raw_expression)")
    print(f"   - 录入标准映射规则: {len(mapping_rows)} 种不同表述 (表: s9_label_mapping)")

    # 打印汇总汇报
    print("\n" + "="*80)
    print("Section 9 已归类进标准标签字段的原文不同表述汇总表")
    print("="*80)
    cur.execute("""
    SELECT std_seq, std_name, raw_label, occurrences_count, models_sample, sample_values
    FROM s9_label_mapping
    ORDER BY CAST(SUBSTR(std_seq, 3) AS INTEGER), std_name, occurrences_count DESC
    """)
    rows = cur.fetchall()
    cur_std = ""
    for seq, std, raw, count, models, sample in rows:
        if std != cur_std:
            print(f"\n### 【{seq} {std}】")
            cur_std = std
        print(f"- 原文表述: `{raw}` | 命中型号数: {count} | 示例值: {sample}")
    print("="*80 + "\n")

    return mapping_stats, raw_rows

if __name__ == "__main__":
    db = r"F:\正式项目与模块化内容\冠志\MSDS\03-数据库\正式库\Data Base\msds_standard.db"
    docs = r"F:\正式项目与模块化内容\冠志\MSDS\03-数据库\正式库\入库word  第一批"
    run_audit(db, docs)
