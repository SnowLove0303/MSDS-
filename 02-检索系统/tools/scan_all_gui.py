# -*- coding: utf-8 -*-
"""在整个冠志项目根目录下搜索所有关于 GUI 目录树、节导航、序号显示的实现."""

import os
from pathlib import Path

ROOT = Path(r"F:\正式项目与模块化内容\冠志")

def scan_all_gui():
    print(f"Scanning {ROOT}...")
    for root, dirs, files in os.walk(ROOT):
        # 排除 git、venv、.pytest_cache、node_modules
        if any(x in root for x in [".git", "venv", ".pytest_cache", "__pycache__", "node_modules", "outputs"]):
            continue
        for f in files:
            if f.endswith(".py") and any(k in f for k in ["gui", "tree", "nav", "search", "main", "db"]):
                p = Path(root) / f
                content = p.read_text(encoding="utf-8", errors="ignore")
                for i, line in enumerate(content.splitlines(), 1):
                    if any(term in line for term in ["tree.insert", "sec_tree", "SEC_TITLES", "sn.title", "sn.number", "SectionTree"]):
                        print(f"{p.relative_to(ROOT)}:{i} -> {line.strip()}")

if __name__ == "__main__":
    scan_all_gui()
