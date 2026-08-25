# -*- coding: utf-8 -*-
"""命令行: 三表数据库构建 (标准汇总表 + 原始表 + 字段映射表)."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from core.database import build_database

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("用法: python tools/build_database.py <入库目录> <输出xlsx>")
        sys.exit(1)
    info = build_database(sys.argv[1], sys.argv[2])
    print(f"OK: {info['files']} 文件 / 标准汇总 {info['cols']}列 "
          f"映射 {info['mappings']} 条 / 原始 {info['raw_rows']} 行")
    if info['failed']:
        print("跳过失败:", info['failed'])
    print("输出:", info['out'])
