# -*- coding: utf-8 -*-
"""数据库检索 API 测试 (SQLite 库: 建库 → 型号检索 → 关键词检索 → 渲染).

覆盖 PRD 验收: 型号唯一索引 (精确/模糊) / 关键词检索 (多词 AND) /
三级树渲染 (S2 序号行独立父级, S11 国标大类归并) / 范围过滤 / 节行渲染.
"""
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from core import msds_db
from core.msds_db import (find_models, listed_section_rows, listed_tree_nodes,
                          model_search, open_db, render_model_json,
                          render_model_tsv, render_model_tree)

TPL = ROOT / "templates" / "MSDS_CN 国彩 模板.docx"


@pytest.fixture()
def conn(tmp_path):
    """临时库: 模板 docx 入库 → (conn, model_id)."""
    path = tmp_path / "t.db"
    c = open_db(str(path))
    mid = msds_db.insert_docx(c, str(TPL), msds_db.wide_columns())
    return c, mid


def test_find_models_exact(conn):
    c, mid = conn
    hits = find_models(c, "PEA-4139")
    assert hits and hits[0][1] == "PEA-4139" and hits[0][0] == mid


def test_find_models_fuzzy(conn):
    c, _ = conn
    hits = find_models(c, "4139")
    assert any(h[1] == "PEA-4139" for h in hits)
    assert find_models(c, "不存在的型号XYZ") == []


def test_tree_s2_flat_parents(conn):
    """S2: 定稿模板结构 — 序号行均为独立父级 (2.1~2.9 平级, 含四子列)."""
    c, mid = conn
    nodes = listed_tree_nodes(c, mid, {2})
    assert len(nodes) == 1
    sn = nodes[0]
    titles = [(b.seq, b.title) for b in sn.big_titles]
    assert ("2.1", "紧急情况概述") in titles
    assert ("2.2", "GHS危险性类别") in titles
    assert ("2.4", "信号词") in titles
    assert ("2.5", "危险性说明") in titles
    assert ("2.5", "防范说明") in titles
    assert ("2.6", "物理和化学危险") in titles
    assert ("2.7", "健康危害") in titles
    assert ("2.8", "环境危害") in titles
    assert ("2.9", "其他危害") in titles
    # 2.3 GHS标签要素 是带值字段, 象形图挂其下
    b23 = next(b for b in sn.big_titles if b.seq == "2.3")
    assert any(c2.label == "GHS象形图" for c2 in b23.children)


def test_tree_s9_seq_parents(conn):
    """S9: 9.1~9.23 各自独立父级且带值 (序号行 = 父级)."""
    c, mid = conn
    nodes = listed_tree_nodes(c, mid, {9})
    sn = nodes[0]
    seqs = [b.seq for b in sn.big_titles]
    assert "9.1" in seqs and "9.23" in seqs
    b91 = next(b for b in sn.big_titles if b.seq == "9.1")
    assert b91.kind == "field" and b91.value  # 二级带值


def test_s11_merged_major(conn):
    """S11: 字段行归并为国标大类 (11.1~11.10), 总结句保留."""
    c, mid = conn
    nodes = listed_tree_nodes(c, mid, {11})
    sn = nodes[0]
    titles = [b.title for b in sn.big_titles]
    assert "急性毒性" in titles
    assert "致突变性" in titles
    assert "附加信息" in titles
    notes = [f for b in sn.big_titles for f in b.children if f.kind == "note"]
    notes += [f for f in sn.direct_fields if f.kind == "note"]
    assert any("毒理学研究" in (f.value or "") for f in notes)


def test_section_rows_s11(conn):
    """listed_section_rows: S11 骨架 — 11 大类 + note 槽位 (产品说明/引导段)."""
    c, mid = conn
    rows = listed_section_rows(c, mid, 11)
    kinds = {r.kind for r in rows}
    assert "note" in kinds
    notes = [r for r in rows if r.kind == "note"]
    types = {r.label for r in notes}
    assert "产品说明" in types or "成分参考引导段" in types


def test_search_kw(conn):
    """关键词检索: 标签命中 + 多词 AND."""
    c, mid = conn
    hits = model_search(c, "闪点")
    assert hits and any(h[2] == 9 and h[3] == "9.6" for h in hits)
    hits2 = model_search(c, "闪点 76")
    assert hits2  # value 含 76℃
    assert model_search(c, "绝无此词zzz") == []


def test_search_sections_filter(conn):
    """范围过滤: --sections 语义, 只命中指定节."""
    c, mid = conn
    hits = model_search(c, "闪点", sections={2})
    assert all(h[2] == 2 for h in hits)


def test_render_tree_json_tsv(conn):
    """三种渲染格式可用且含型号内容."""
    c, mid = conn
    tree = render_model_tree(c, mid)
    assert "GHS危险性类别" in tree
    js = render_model_json(c, mid)
    assert '"section": 9' in js
    tsv = render_model_tsv(c, mid)
    assert tsv.splitlines()[0].startswith("文件\t节\t")
    assert "PEA-4139" in tsv


def test_section_rows_subtable(tmp_path):
    """subtable 行 (S8 生物限值): EC-1801 有子表 → 行携带 sub_header/sub_rows."""
    ec = ROOT.parent / "03-数据库" / "正式库" / "入库word  第一批" \
        / "EC-1801 msds_CN 冠志_模板覆写输出.docx"
    if not ec.exists():
        pytest.skip("EC-1801 样例文件不存在")
    path = tmp_path / "t2.db"
    c = open_db(str(path))
    mid = msds_db.insert_docx(c, str(ec), msds_db.wide_columns())
    rows = listed_section_rows(c, mid, 8)
    sub = [r for r in rows if r.kind == "subtable"]
    assert sub
    assert sub[0].sub_header and "组分名称" in sub[0].sub_header


def test_search_s1_s3_s9_enhanced(conn):
    """S1/S3/S9 检索增强: 序号检索、别名检索、跨型号属性组合检索."""
    c, mid = conn
    # 序号检索
    assert model_search(c, "9.6")  # 9.6 闪点
    assert model_search(c, "1.2")  # 1.2 使用建议
    assert model_search(c, "3.1")  # 3.1 产品类型

    # 别名/同义词检索 (schema_field 别名)
    assert model_search(c, "Flash point")
    assert model_search(c, "用途")
    assert model_search(c, "Mixtures")

    # 型号 + 属性组合检索 (多词 AND)
    assert model_search(c, "PEA-4139 闪点")
    assert model_search(c, "PEA-4139 供应商")


def test_s3_component_standardization(conn):
    """测试 Section 3 成分名称标准化收敛与检索呈现."""
    from core.s3_component_std import standardize_component_name
    # 别名/异写收敛验证
    assert standardize_component_name("N-乙基吡咯烷酮", "2687-91-4") == "N-乙基吡咯烷酮 (NEP)"
    assert standardize_component_name("N-乙基吡咯烷酮(NEP)") == "N-乙基吡咯烷酮 (NEP)"
    assert standardize_component_name("苯并异噻唑啉酮(BIT)", "2634-33-5") == "1,2-苯并异噻唑-3-酮 (BIT)"
    assert standardize_component_name("N,N二甲基乙醇胺", "108-01-0") == "N,N-二甲基乙醇胺 (DMEA)"
    assert standardize_component_name("丙二醇丁醚", "5131-66-8") == "1-丁氧基-2-丙醇 (PnB / 丙二醇丁醚)"

