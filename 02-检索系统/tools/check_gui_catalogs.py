# -*- coding: utf-8 -*-
"""检查 GUI 目录树与节点标题中的序号格式与重复问题."""

import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from core.msds_db import SEC_TITLES, listed_tree_nodes, open_db, model_tree_nodes

DB_PATH = r"F:\正式项目与模块化内容\冠志\MSDS\03-数据库\正式库\Data Base\msds_standard.db"

def inspect_gui_catalogs():
    conn = open_db(DB_PATH)
    mid = conn.execute("SELECT model_id FROM msds_model LIMIT 1").fetchone()[0]

    print("=== 1. SEC_TITLES 字典中的标题 ===")
    for k, v in sorted(SEC_TITLES.items()):
        print(f"  [{k}] -> \"{v}\"")

    print("\n=== 2. listed_tree_nodes 产生的 SectionNode ===")
    nodes = listed_tree_nodes(conn, mid)
    for sn in nodes:
        fmt_tree = f"{sn.number}. {sn.title}"
        print(f"  sn.number={sn.number:<2} | sn.title=\"{sn.title}\" | 树节点格式=\"{fmt_tree}\"")
        for bt in sn.big_titles[:3]:
            print(f"     └─ bt.seq=\"{bt.seq}\", bt.title=\"{bt.title}\"")

    print("\n=== 3. model_tree_nodes 产生的 SectionNode ===")
    mnodes = model_tree_nodes(conn, mid)
    for sn in mnodes:
        fmt_tree = f"{sn.number}. {sn.title}"
        print(f"  sn.number={sn.number:<2} | sn.title=\"{sn.title}\" | 树节点格式=\"{fmt_tree}\"")

if __name__ == "__main__":
    inspect_gui_catalogs()
