# -*- coding: utf-8 -*-
"""Section 3 全量成分与 CAS 检索、聚类归一、去重分析脚本."""

import sqlite3
import re
import sys
from pathlib import Path
from collections import defaultdict, Counter

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from core.docx_reader import read_msds
from core.msds_db import open_db

DB_PATH = r"F:\正式项目与模块化内容\冠志\MSDS\03-数据库\正式库\Data Base\msds_standard.db"
DST_DIR = Path(r"F:\正式项目与模块化内容\冠志\MSDS\03-数据库\正式库\入库word  第一批")

def analyze_all_s3_components():
    conn = open_db(DB_PATH)
    cur = conn.cursor()

    cur.execute("""
    SELECT m.model, f.label, f.value
    FROM msds_field f
    JOIN msds_model m ON f.model_id = m.model_id
    WHERE f.section = 3 AND f.kind = 'component'
    ORDER BY f.model_id, f.row_index
    """)
    db_rows = cur.fetchall()

    # 1. 结构化解析
    cas_to_info = defaultdict(lambda: {"names": Counter(), "models": set(), "sample_vals": []})
    non_cas_to_info = defaultdict(lambda: {"names": Counter(), "models": set(), "cas_types": Counter(), "sample_vals": []})

    for model, raw_name, val in db_rows:
        raw_name = raw_name.strip() if raw_name else ""
        if not raw_name:
            continue
        
        # 从 val 中提取 CAS: xxx
        cas_m = re.search(r"CAS[:：\s]*([^\s|]+)", val)
        raw_cas = cas_m.group(1).strip() if cas_m else ""
        
        # 纯化 CAS 格式
        clean_cas = re.sub(r"[^\d\-]", "", raw_cas)
        if re.match(r"^\d{2,7}-\d{2}-\d$", clean_cas):
            cas_to_info[clean_cas]["names"][raw_name] += 1
            cas_to_info[clean_cas]["models"].add(model)
            if len(cas_to_info[clean_cas]["sample_vals"]) < 3:
                cas_to_info[clean_cas]["sample_vals"].append(val)
        else:
            # 归一化无 CAS 或保密成分名
            norm_name = re.sub(r"[\s\t\r\n]", "", raw_name)
            cas_type = "商业机密" if ("机密" in raw_cas or "保密" in raw_cas or "机密" in val) else "未提供/无CAS"
            non_cas_to_info[norm_name]["names"][raw_name] += 1
            non_cas_to_info[norm_name]["models"].add(model)
            non_cas_to_info[norm_name]["cas_types"][cas_type] += 1
            if len(non_cas_to_info[norm_name]["sample_vals"]) < 3:
                non_cas_to_info[norm_name]["sample_vals"].append(val)

    print("=" * 80)
    print(f"【Section 3 全量成分与 CAS 聚类去重统计】")
    print(f"总计提取成分明细记录: {len(db_rows)} 条")
    print(f"聚类独立有效 CAS 数量: {len(cas_to_info)} 个")
    print(f"聚类独立高分子树脂/保密/无CAS成分分类: {len(non_cas_to_info)} 个")
    print("=" * 80)

    print("\n--- 1. 具有确定 CAS 号的成分清单 (去重聚类) ---")
    for cas, info in sorted(cas_to_info.items(), key=lambda x: -len(x[1]["models"])):
        primary_name, _ = info["names"].most_common(1)[0]
        aliases = [f"「{k}」({v}次)" for k, v in info["names"].most_common() if k != primary_name]
        alias_str = " | 异写/同义: " + ", ".join(aliases) if aliases else ""
        print(f"CAS: {cas:<12} | 主成分名: {primary_name:<25} | 覆盖型号: {len(info['models']):<3} {alias_str}")

    print("\n--- 2. 高分子树脂/保密/无独立CAS成分清单 (去重聚类) ---")
    for norm_name, info in sorted(non_cas_to_info.items(), key=lambda x: -len(x[1]["models"])):
        primary_name, _ = info["names"].most_common(1)[0]
        aliases = [f"「{k}」({v}次)" for k, v in info["names"].most_common() if k != primary_name]
        alias_str = " | 异写/同义: " + ", ".join(aliases) if aliases else ""
        cas_type = info["cas_types"].most_common(1)[0][0]
        print(f"成分: {primary_name:<30} | CAS状态: {cas_type:<10} | 覆盖型号: {len(info['models']):<3} {alias_str}")

if __name__ == "__main__":
    analyze_all_s3_components()
