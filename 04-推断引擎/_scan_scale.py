# -*- coding: utf-8 -*-
"""临时: 统计 429 份 MSDS 反推成分的规模 (决定种子表 vs 数据库)."""
import sys, hashlib, re
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "结构读取"))
from core.docx_reader import read_msds

BASES = [ROOT / "数据库/MSDS/中文", ROOT / "数据库/MSDS/英文", ROOT / "数据库/MSDS/补充库"]

def sha1(p):
    h = hashlib.sha1()
    h.update(p.read_bytes())
    return h.hexdigest()

seen = set()
comp_names = Counter()
cas_map = defaultdict(set)      # cas -> 出现产品
conc_ok = 0
total_comps = 0
no_cas = 0

files = []
for b in BASES:
    if b.exists():
        files.extend(b.rglob("*.docx"))
for p in files:
    if p.name.startswith("~$"):
        continue
    h = sha1(p)
    if h in seen:
        continue
    seen.add(h)
    try:
        r = read_msds(p)
    except Exception:
        continue
    s3 = r.section(3)
    comps = s3.components if s3 else []
    for c in comps:
        total_comps += 1
        name = (c.name or "").strip()
        cas = (c.cas or "").strip().upper()
        conc = (c.conc or "").strip()
        if name:
            comp_names[name] += 1
        if cas and re.fullmatch(r"\d{2,7}-\d{2}-\d", cas):
            cas_map[cas].add(p.stem)
        else:
            no_cas += 1
        if conc and conc not in ("商业机密", "不适用", "无数据", "-"):
            conc_ok += 1

real_cas = {c for c, prods in cas_map.items() if prods}
print(f"去重文件: {len(seen)}")
print(f"成分总行数: {total_comps}  (有效含量 {conc_ok})")
print(f"唯一成分名: {len(comp_names)}")
print(f"唯一合法CAS: {len(real_cas)}  (无有效CAS {no_cas})")
print(f"Top 20 成分名:")
for name, n in comp_names.most_common(20):
    print(f"  {name}: {n}")
print(f"CAS出现产品数分布:")
dist = Counter(len(v) for v in cas_map.values())
for k in sorted(dist):
    print(f"  {k} 个产品: {dist[k]} 个CAS")
print(f"Top CAS (按产品数):")
for cas, prods in sorted(cas_map.items(), key=lambda kv: -len(kv[1]))[:8]:
    print(f"  {cas}: {len(prods)} 个产品 | 例: {sorted(prods)[:3]}")
