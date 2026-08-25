# -*- coding: utf-8 -*-
"""标签归一化审计 (audit_labels): 扫描 docx 提取所有 (节,标签) 的 Schema 命中情况.

输出:
  1. 未命中清单 (节/标签/出现次数/示例值) — 用于数据驱动补 Schema
  2. 命中率统计

用法:
  python tools/audit_labels.py <docx或目录...>
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from collections import Counter, defaultdict

from core.docx_reader import read_msds
from core.schema import standard_fields, standard_name


def collect(inputs: list[str]):
    files = []
    for raw in inputs:
        p = Path(raw)
        if p.is_dir():
            files.extend(sorted(f for f in p.glob("*.docx")
                                if not f.name.startswith("~$")))
        elif p.exists() and p.suffix.lower() == ".docx":
            files.append(p)
    return sorted(set(files), key=lambda x: str(x).lower())


def main():
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        return 2
    files = collect(args)
    print(f"扫描 {len(files)} 份 docx ...")

    # (section, label) -> 命中? ; 未命中详情
    hit_cnt = Counter()
    miss: dict[tuple[int, str], dict] = defaultdict(
        lambda: {"count": 0, "example": ""})

    for p in files:
        try:
            r = read_msds(str(p))
        except Exception as exc:
            print(f"  ✗ {p.name}: {exc}")
            continue
        for num, sec in sorted(r.sections.items()):
            for row in sec.iter_rows():
                if row.kind not in ("field", "sub"):
                    continue
                label = (row.label or "").strip()
                if not label:
                    continue
                std = standard_name(num, label)
                # 命中 = 该标准名在 Schema 中确有定义 (自身名或别名)
                defined = any(f.name == std for f in standard_fields(num))
                if defined:
                    hit_cnt[num] += 1
                else:
                    key = (num, label)
                    miss[key]["count"] += 1
                    if not miss[key]["example"]:
                        miss[key]["example"] = (row.value or "")[:40].replace("\n", " ")

    total_hit = sum(hit_cnt.values())
    total_miss = sum(v["count"] for v in miss.values())
    print(f"\n命中(标准字段): {total_hit} | 未命中: {total_miss} | 未命中标签种类: {len(miss)}")
    print(f"\n=== 未命中清单 (按节) ===")
    cur = None
    for (num, label), info in sorted(miss.items()):
        if num != cur:
            cur = num
            print(f"\n--- S{num} ---")
        print(f"  {label:<28} ×{info['count']:<3} 例: {info['example'][:40]}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
