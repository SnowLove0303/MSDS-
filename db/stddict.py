# -*- coding: utf-8 -*-
"""标准字段字典半自动构建 (P2).

核心规则 (用户确立):
  - 意思一致 → 归同一标准字段, 列统一一种表达
  - 组内采用"最多次使用的"表达作标准名 (频次裁决)
  - 原有表达不丢 → 存入 aliases (field_dict.aliases_json)
  - 未归并标签 → status='pending' 待确认池 (宁缺毋滥, 不丢)

流程:
  1. seed 同义组: 静态 _NORM (人工确认过) + 已有 field_dict active 条目
  2. 扫描文档: 按节统计原始标签频次 + 值样例
  3. 归一 key 相同 → 自动归组 (格式差异收敛)
  4. 频次裁决: 组内频次最高表达 → std_field, 其余 → aliases
  5. 未命中任何组的新标签 → 孤立组, status='pending'
  6. suggest_merges: 对 pending 给同节归一化核最接近的 active 目标建议 (供 GUI 确认)
"""
from __future__ import annotations

import difflib
import json
import sys
from collections import Counter, defaultdict
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))
if str(_ROOT / "批量化读取") not in sys.path:
    sys.path.insert(0, str(_ROOT / "批量化读取"))

from build_db import _NORM, _nk, normalize_field

from .catalog import strip_label_prefix


# ============================================================
# 扫描: 按节统计标签频次 + 值样例
# ============================================================

def scan_db(conn, progress=None):
    """从总库 fields 表统计 (与 scan_files 同结构, 不重读源文件).

    用 raw_label 原始表达; value 非空时取作样例.
    """
    stats: dict = {}
    n = 0
    for r in conn.execute(
            "SELECT section, raw_label, value FROM fields "
            "WHERE kind='field' AND raw_label!='' ORDER BY section, row_order"):
        n += 1
        sec = r["section"]
        if sec == 0:   # S0 页眉页脚独立管理, 不参与字段归并
            continue
        lab = strip_label_prefix(r["raw_label"])
        nk = _nk(lab)
        if not nk:
            continue
        sec = r["section"]
        key = (sec, nk)
        st = stats.setdefault(key, {
            "freq": 0, "labels": Counter(), "samples": [],
            "section": sec, "static_hit": False})
        st["freq"] += 1
        st["labels"][lab] += 1
        v = (r["value"] or "").strip()
        if v and v not in st["samples"]:
            st["samples"].append(v)
        if len(st["samples"]) > 3:
            st["samples"] = st["samples"][:3]
        if _nk(normalize_field(lab)) == nk:
            st["static_hit"] = True
        if progress and n % 2000 == 0:
            progress(f"已统计 {n} 行")
    return stats


def scan_files(files, progress=None):
    """扫描 docx, 统计 {(section, nk_label): 统计}.

    统计内容: freq(文档数), labels(原始表达 Counter), samples(值样例, 去重≤3),
    static_hit(静态词典是否覆盖), section.
    """
    from core.docx_reader import read_msds

    stats: dict = {}
    for idx, p in enumerate(files, 1):
        try:
            r = read_msds(p)
        except Exception as exc:
            if progress:
                progress(f"[{idx}] 读取失败 {p.name}: {exc}")
            continue
        for sec_num, sec in r.sections.items():
            if sec_num == 0:   # S0 页眉页脚独立管理, 不参与字段归并
                continue
            for row in sec.iter_rows():
                if row.kind != "field" or not row.label:
                    continue
                lab = strip_label_prefix(row.label)
                nk = _nk(lab)
                if not nk:
                    continue
                key = (sec_num, nk)
                st = stats.setdefault(key, {
                    "freq": 0, "labels": Counter(), "samples": [],
                    "section": sec_num, "static_hit": False})
                st["freq"] += 1
                st["labels"][lab] += 1
                v = (row.value or "").strip()
                if v and v not in st["samples"]:
                    st["samples"].append(v)
                if len(st["samples"]) > 3:
                    st["samples"] = st["samples"][:3]
                if _nk(normalize_field(lab)) == nk:
                    st["static_hit"] = True
        if progress and idx % 50 == 0:
            progress(f"已扫描 {idx}/{len(files)}")
    return stats


# ============================================================
# 同义组 + 频次裁决
# ============================================================

def _seed_groups(existing_active=None):
    """构建同义组: {nk_alias: group_key}. 组源 = 静态 _NORM + 已有 active 条目.

    group_key = 静态标准名 (供组标识); 实际标准名由频次裁决决定.
    """
    groups: dict[str, str] = {}
    for std, aliases in _NORM.items():
        for a in aliases:
            groups[_nk(a)] = std
    if existing_active:
        for std, aliases in existing_active:
            groups[_nk(std)] = std
            for a in aliases:
                groups[_nk(a)] = std
    return groups


def build_field_dict(conn, files=None, stats=None, progress=None, min_freq=1) -> dict:
    """扫描 → 归组 → 频次裁决 → 写 field_dict.

    stats 可由 scan_files(docx) 或 scan_db(总库) 提供; 均缺省时走 scan_files.
    返回 stats: {groups, static_groups, merged, pending, freq_rerouted, unmatched_docs}.
    """
    from .catalog import get_field_keys  # noqa

    existing_active = [(r["std_field"], json.loads(r["aliases_json"] or "[]"))
                       for r in conn.execute(
                           "SELECT std_field, aliases_json FROM field_dict "
                           "WHERE status='active'")]
    seed = _seed_groups(existing_active)

    # 重建语义: 清空旧 pending (重算), active 保留作同义种子 (仍会按频次覆盖)
    conn.execute("DELETE FROM field_dict WHERE status='pending'")

    if stats is None:
        stats = scan_files(files, progress)
    # 按 group_key 聚合: {group_key: {section, labels: Counter, samples, freq}}
    groups: dict = {}
    orphan = {}
    for (sec, nk), st in stats.items():
        g = seed.get(nk)
        if g:
            grp = groups.setdefault((g, sec), {
                "labels": Counter(), "samples": set(), "freq": 0, "section": sec})
        else:
            grp = orphan.setdefault((nk, sec), {
                "labels": Counter(), "samples": set(), "freq": 0, "section": sec,
                "group_key": nk})
        grp["freq"] += st["freq"]
        grp["labels"].update(st["labels"])
        grp["samples"].update(st["samples"])

    n_groups = len(groups)
    n_orphan = len(orphan)
    n_rerouted = 0

    # 写 active 组 (频次裁决: 组内频次最高 → std_field)
    for (g, sec), grp in groups.items():
        labels = grp["labels"]
        if not labels:
            continue
        top_label, top_freq = labels.most_common(1)[0]
        std = top_label
        aliases = [lb for lb, _ in labels.most_common()]
        aliases.remove(std)
        # 若原静态标准名也在组内且频次最高, 则用静态标准名 (更规范)
        if g in labels and labels[g] == top_freq:
            std = g
        if std != g and g in labels:
            n_rerouted += 1
        conn.execute(
            "INSERT INTO field_dict (std_field, section, display_order, freq, "
            "aliases_json, status) VALUES (?,?,?,?,?, 'active') "
            "ON CONFLICT(std_field, section) DO UPDATE SET "
            "freq=excluded.freq, aliases_json=excluded.aliases_json",
            (std, sec, grp["freq"], grp["freq"],
             json.dumps(aliases, ensure_ascii=False)))
    conn.commit()

    # 写 pending 组 (孤立标签: 未命中任何种子, 但归一 key 相同已自动合并)
    for (nk, sec), grp in orphan.items():
        labels = grp["labels"]
        std = labels.most_common(1)[0][0]
        aliases = [lb for lb, _ in labels.most_common()]
        aliases.remove(std)
        conn.execute(
            "INSERT INTO field_dict (std_field, section, display_order, freq, "
            "aliases_json, status) VALUES (?,?,?,?,?, 'pending')",
            (std, sec, grp["freq"], grp["freq"], json.dumps(aliases, ensure_ascii=False)))
    conn.commit()

    return {
        "groups": n_groups, "pending": n_orphan, "freq_rerouted": n_rerouted,
        "stats_count": len(stats),
    }


# ============================================================
# 待确认池: 建议合并目标 (半自动人工裁决辅助)
# ============================================================

def _label_similarity(a: str, b: str) -> float:
    """归一化标签相似度 (difflib). 0..1."""
    return difflib.SequenceMatcher(None, _nk(a), _nk(b)).ratio()


def suggest_merges(conn, top_k=3, min_sim=0.5):
    """对 pending 标签, 同节内找归一化核最接近的 active 标准字段, 给建议.

    返回: [{'section', 'pending', 'pending_freq', 'candidates': [
             {'target', 'target_freq', 'similarity', 'pending_samples', 'target_samples'}]}]
    供 GUI 一键采纳: 把 pending 的 aliases 并入 target, 置 active.
    """
    pendings = [dict(r) for r in conn.execute(
        "SELECT id, std_field, section, freq, aliases_json FROM field_dict "
        "WHERE status='pending' ORDER BY section, freq DESC")]
    actives = [dict(r) for r in conn.execute(
        "SELECT std_field, section, freq FROM field_dict WHERE status='active'")]
    by_sec: dict[int, list[dict]] = defaultdict(list)
    for a in actives:
        by_sec[a["section"]].append(a)

    out = []
    for p in pendings:
        cands = []
        for a in by_sec.get(p["section"], []):
            sim = _label_similarity(p["std_field"], a["std_field"])
            if sim >= min_sim:
                cands.append({"target": a["std_field"], "target_freq": a["freq"],
                              "similarity": round(sim, 2)})
        if not cands:
            # 跨节: 找同 std_field 名 (归一化相同) 的 active 兜底
            for a in actives:
                if _nk(a["std_field"]) == _nk(p["std_field"]):
                    cands.append({"target": a["std_field"], "target_freq": a["freq"],
                                  "similarity": 1.0})
        cands.sort(key=lambda x: (-x["similarity"], -x["target_freq"]))
        if cands:
            out.append({
                "section": p["section"], "pending": p["std_field"],
                "pending_freq": p["freq"], "candidates": cands[:top_k],
            })
    return out


def merge_pending(conn, pending_std: str, section: int, target_std: str) -> bool:
    """GUI 采纳: 把 pending 标签并入 target 标准字段 (aliases 合并, 置 active).

    返回是否成功 (target 不存在 → False).
    """
    t = conn.execute("SELECT id, aliases_json FROM field_dict "
                     "WHERE std_field=? AND section=? AND status='active'",
                     (target_std, section)).fetchone()
    if not t:
        return False
    aliases = json.loads(t["aliases_json"] or "[]")
    if pending_std not in aliases:
        aliases.append(pending_std)
    conn.execute("UPDATE field_dict SET aliases_json=?, freq=freq+? WHERE id=?",
                 (json.dumps(aliases, ensure_ascii=False), pending_freq(conn, pending_std, section), t["id"]))
    conn.execute("DELETE FROM field_dict WHERE std_field=? AND section=? AND status='pending'",
                 (pending_std, section))
    conn.commit()
    return True


def pending_freq(conn, std: str, section: int) -> int:
    r = conn.execute("SELECT freq FROM field_dict WHERE std_field=? AND section=?",
                     (std, section)).fetchone()
    return r["freq"] if r else 0


def set_pending_active(conn, std: str, section: int) -> None:
    """直接把孤立 pending 标签升为 active (确认它是独立标准字段)."""
    conn.execute("UPDATE field_dict SET status='active' WHERE std_field=? AND section=?",
                 (std, section))
    conn.commit()


def list_pending(conn) -> list[dict]:
    return [dict(r) for r in conn.execute(
        "SELECT * FROM field_dict WHERE status='pending' ORDER BY section, freq DESC")]
