# -*- coding: utf-8 -*-
"""深度排查 GUI 和 CLI 中所有节、子标题、字段中的序号重复问题."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from core.docx_reader import read_msds
from core.msds_db import open_db, listed_tree_nodes, SEC_TITLES
from core.extract import build_hierarchy, render_tree

DB_PATH = r"F:\正式项目与模块化内容\冠志\MSDS\03-数据库\正式库\Data Base\msds_standard.db"
DST_DIR = Path(r"F:\正式项目与模块化内容\冠志\MSDS\03-数据库\正式库\入库word  第一批")

def check_all_number_duplications():
    conn = open_db(DB_PATH)
    mid = conn.execute("SELECT model_id FROM msds_model LIMIT 1").fetchone()[0]

    print("=== 1. 检查 SEC_TITLES 与 SectionNode 节标题 ===")
    for k, v in sorted(SEC_TITLES.items()):
        print(f"  SEC_TITLES[{k}] = '{v}'")

    print("\n=== 2. 检查 listed_tree_nodes 中的所有节点标题与序号 ===")
    nodes = listed_tree_nodes(conn, mid)
    for sn in nodes:
        print(f"Section {sn.number}: title='{sn.title}', full_title='{sn.full_title}'")
        for bt in sn.big_titles:
            # 检查是否有形如 "1.1 1.1产品名称" 或 "1.1. 1.1."
            if bt.seq and bt.seq in bt.title:
                print(f"  [发现子标题序号重复] seq='{bt.seq}', title='{bt.title}'")
            for fn in bt.children:
                if fn.label and any(f"{sn.number}." in fn.label for _ in [1]):
                    print(f"  [发现字段标签含节号] label='{fn.label}'")

    print("\n=== 3. 检查从 docx 原文档解析的 ParseResult ===")
    sample_file = next(DST_DIR.glob("*.docx"))
    print(f"Sample docx: {sample_file.name}")
    res = read_msds(sample_file)
    for n, sec in sorted(res.sections.items()):
        print(f"Docx Section {n}: title='{sec.title}', full_title='{sec.full_title}'")
        for row in sec.iter_rows():
            if row.kind == "sub" and row.seq and row.seq in row.label:
                print(f"  [Docx sub 序号重复] seq='{row.seq}', label='{row.label}'")

if __name__ == "__main__":
    check_all_number_duplications()
