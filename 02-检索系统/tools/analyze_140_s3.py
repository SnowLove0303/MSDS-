# -*- coding: utf-8 -*-
"""深度排查 Section 3 所有 140 种原文候选表述并聚类."""

import re
import sys
from pathlib import Path
from collections import defaultdict

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from tools.deep_scan_s3 import deep_scan_word_s3

def analyze_all_raw_candidates():
    raw_cas, raw_comps = deep_scan_word_s3()

    print(f"=== 深度排查 140 种原文成分表述 ===")
    
    # 按照是否有 CAS 分类
    with_cas = {}
    without_cas = {}

    for comp_name, info in raw_comps.items():
        cas_list = list(info["cas_list"])
        if cas_list:
            with_cas[comp_name] = (cas_list, len(info["models"]), info["models"])
        else:
            without_cas[comp_name] = (len(info["models"]), info["models"])

    print(f"\n1. 具有直接关联 CAS 号的成分表述 ({len(with_cas)} 种):")
    for name, (cas, m_cnt, models) in sorted(with_cas.items(), key=lambda x: -x[1][1]):
        print(f"  • {name:<35} | CAS: {', '.join(cas):<25} | 涉及型号数: {m_cnt}")

    print(f"\n2. 无独立 CAS 号 / 商业机密 / 聚合物表述 ({len(without_cas)} 种):")
    for name, (m_cnt, models) in sorted(without_cas.items(), key=lambda x: -x[1][0]):
        print(f"  • {name:<35} | 涉及型号数: {m_cnt}")

if __name__ == "__main__":
    analyze_all_raw_candidates()
