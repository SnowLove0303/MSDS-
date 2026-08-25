# -*- coding: utf-8 -*-
"""去重合并逻辑测试: 查重范围 = lines block 内 field 完全重复.

背景: 不同 MSDS 写法不一, 排版缺陷导致同一字段块在节内重复出现
(如 OS-1330 S2 的 防范说明 + 2.3 其他危险 整段复制两次). 去重逻辑
只处理 lines block 内解析出的 field 行**完全重复** (seq+label+value
全同), 保留首次. 查重范围刻意限定在 block 内, 不扩展到节级别 —
节内大量"分隔重复"是合理子项 (如 S11 各毒理测试共用 聚氨酯分散体/
物种 标签), 去重会误删.
"""
from collections import Counter

from core.structure import SectionRow, split_text_block


def _fields(rows: list[SectionRow]) -> list[SectionRow]:
    return [r for r in rows if r.kind == "field"]


def test_dedupe_os1330_style_block():
    """OS-1330 S2 风格: 整段 防范说明+2.3 重复 → 去重为 1 份."""
    block = (
        "防范说明:\nP273 禁止排入环境。\nP501 将本品或其容器送至有资质的废物处理厂处置。\n"
        "2.3 其他危险 无适用资料。\n"
        "防范说明:\nP273 禁止排入环境。\nP501 将本品或其容器送至有资质的废物处理厂处置。\n"
        "2.3 其他危险 无适用资料。"
    )
    rows = split_text_block(block)
    fields = _fields(rows)
    cnt = Counter((r.seq, r.label) for r in fields)
    assert cnt[("", "防范说明")] == 1, f"防范说明应去重为1, 实得 {cnt[('', '防范说明')]}"
    assert cnt[("2.3", "其他危险")] == 1, "2.3 其他危险 应去重为1"
    # 防范说明值含两个 P 代码 (合并保留)
    pf = [r for r in fields if r.label == "防范说明"][0]
    assert "P273" in pf.value and "P501" in pf.value


def test_dedupe_keeps_scattered_subitems():
    """S11 风格: 分隔重复的子项 (聚氨酯分散体/物种) 不去重."""
    block = (
        "急性毒性:\n方法: OECD 423\n聚氨酯分散体\n物种: 大鼠\n分类: 无\n"
        "皮肤刺激:\n方法: OECD 431\n聚氨酯分散体\n物种: 家兔\n分类: 无刺激"
    )
    rows = split_text_block(block)
    fields = _fields(rows)
    cnt = Counter(r.label for r in fields)
    assert cnt["聚氨酯分散体"] == 2, "分隔的子项重复应保留 (查重范围=block 内相邻)"
    assert cnt["物种"] == 2


def test_dedupe_keeps_different_values():
    """同名 field 但值不同 → 不去重 (可能有意义)."""
    block = "方法: OECD 423\n方法: OECD 431"
    rows = split_text_block(block)
    fields = _fields(rows)
    cnt = Counter((r.label, r.value) for r in fields)
    assert cnt[("方法", "OECD 423")] == 1
    assert cnt[("方法", "OECD 431")] == 1
    assert len(fields) == 2


def test_seq_label_value_three_col():
    """2.3 其他危险 无适用资料。 → 序号2.3 | 标签其他危险 | 值无适用资料 三列."""
    rows = split_text_block("2.3 其他危险 无适用资料。")
    fields = _fields(rows)
    assert len(fields) == 1
    r = fields[0]
    assert r.seq == "2.3"
    assert r.label == "其他危险"
    assert r.value == "无适用资料"
    assert r.kind == "field"


def test_dedupe_no_pairs_when_no_dup():
    """无重复的 block → 字段数不变."""
    block = "防范说明:\nP273\n危害性说明:\nH412"
    rows = split_text_block(block)
    fields = _fields(rows)
    assert len(fields) == 2


def test_real_undeduped_os1330_s2():
    """真实未去重 OS-1330 文件: 整段 防范说明+2.3 重复 → 去重为 1 份."""
    from pathlib import Path
    from core.docx_reader import read_msds

    p = Path(r"F:\冠志工作空间\产品\TDS MSDS\TDS MSDS\产品 TDS MSDS -- WORD版本 - 副本\7 水性助剂 OS等\OS-1330 msds_CN 冠志.docx")
    if not p.exists():
        return  # 环境无该文件则跳过
    r = read_msds(p)
    sec = r.sections.get(2)
    assert sec is not None
    cnt = Counter((x.seq, x.label) for x in sec.iter_rows() if x.kind == "field")
    assert cnt[("", "防范说明")] == 1, "防范说明应去重为1份"
    assert cnt[("2.3", "其他危险")] == 1, "2.3 其他危险 应去重为1份"
    # 三列拆分
    rows = [x for x in sec.iter_rows() if x.kind == "field"
            and x.label == "其他危险"]
    assert rows and rows[0].seq == "2.3" and rows[0].value == "无适用资料"
    # 1-丁氧基乙醇 是跨行 note (非独立标签)
    assert any(x.kind == "note" and "1-丁氧基乙醇" in x.value
               for x in sec.iter_rows())
