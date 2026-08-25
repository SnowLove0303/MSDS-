# -*- coding: utf-8 -*-
"""复现并排查 GUI 目录树中的序号重复现象."""

import tkinter as tk
from tkinter import ttk
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from core.docx_reader import TEMPLATE_PATH, read_msds
from core.msds_db import open_db, listed_tree_nodes_from_result, SEC_TITLES

DB_PATH = r"F:\正式项目与模块化内容\冠志\MSDS\03-数据库\正式库\Data Base\msds_standard.db"

def reproduce_duplication():
    print("=== 排查 1: 主窗口左侧目录树 SectionTree 加载骨架节点时的显示 ===")
    res = read_msds(TEMPLATE_PATH)
    nodes = listed_tree_nodes_from_result(res)
    for sn in nodes:
        raw_insert_text = f"{sn.number}. {sn.title}"
        print(f"  [SectionTree 显示文本] -> '{raw_insert_text}' (sn.number={sn.number}, sn.title='{sn.title}')")

    print("\n=== 排查 2: 数据库检索窗口 DbSearchWindow 左侧节导航的显示 ===")
    conn = open_db(DB_PATH)
    for num, title in sorted(SEC_TITLES.items()):
        print(f"  [DbSearchWindow 节导航文本] -> '{title}'")

if __name__ == "__main__":
    reproduce_duplication()
