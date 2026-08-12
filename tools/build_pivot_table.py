# -*- coding: utf-8 -*-
"""命令行生成 MSDS 入库总表 (透视结构), 复用 core.pivot_table.build_pivot_table.

用法:
  python tools/build_pivot_table.py <入库目录> <输出xlsx>
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from core.pivot_table import build_pivot_table  # noqa: E402


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print(__doc__)
        raise SystemExit(1)
    info = build_pivot_table(sys.argv[1], sys.argv[2])
    print(f"✓ 已生成: {info['out']}")
    print(f"  文件数: {info['files']}  列数(含型号): {info['cols']}  "
          f"行数: {info['rows'] + 2}  节数: {info['sections']}")
    if info["failed"]:
        print(f"  ⚠️ 读取失败已跳过 {len(info['failed'])} 个:")
        for f in info["failed"]:
            print(f"    - {f}")
