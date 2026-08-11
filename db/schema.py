# -*- coding: utf-8 -*-
"""MSDS 总库 SQLite schema — 明细 EAV 完整入库 + 标准字段字典.

存储形态:
  - products         二级型号 (含一级类目归属, 软删除 active, sha256 去重)
  - categories       一级类目 (PU / 固化剂 / 稀释剂 ...)
  - fields           字段明细 (双轨: field_key 标准 + raw_label 原始, 全量行入库)
  - components       成分表
  - pictograms       象形图原图 (blob)
  - field_dict       标准字段字典 (std_field + aliases, P2 半自动生成)
  - revisions        版本历史 (更新/删除前快照)

原则 (用户确立):
  1. 所有 MSDS 内容 (节/标题/字段/说明/成分/象形图) 完整入库, 禁止缺失
  2. 写入按字典收敛到标准字段, 读取用 raw_label 原始表达回放
  3. 宽表不双写, 由明细物化生成 (materialize)
"""
from __future__ import annotations

SCHEMA = """
CREATE TABLE IF NOT EXISTS categories (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE,
    sort_order INTEGER DEFAULT 0,
    note TEXT DEFAULT ''
);

CREATE TABLE IF NOT EXISTS products (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    category_id INTEGER REFERENCES categories(id),
    model_name TEXT NOT NULL,
    file_name TEXT DEFAULT '',
    file_path TEXT DEFAULT '',
    sha256 TEXT DEFAULT '',
    language TEXT DEFAULT '',
    template_vendor TEXT DEFAULT '',
    product_name TEXT DEFAULT '',
    version TEXT DEFAULT '',
    revision_date TEXT DEFAULT '',
    active INTEGER DEFAULT 1,
    created_at TEXT DEFAULT '',
    updated_at TEXT DEFAULT '',
    note TEXT DEFAULT '',
    UNIQUE(category_id, model_name)
);
CREATE INDEX IF NOT EXISTS idx_prod_active ON products(active);
CREATE INDEX IF NOT EXISTS idx_prod_sha ON products(sha256);
CREATE INDEX IF NOT EXISTS idx_prod_model ON products(model_name);

CREATE TABLE IF NOT EXISTS fields (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    product_id INTEGER NOT NULL REFERENCES products(id) ON DELETE CASCADE,
    section INTEGER NOT NULL,
    seq TEXT DEFAULT '',
    field_key TEXT DEFAULT '',
    raw_label TEXT DEFAULT '',
    value TEXT DEFAULT '',
    kind TEXT DEFAULT 'field',
    editable INTEGER DEFAULT 1,
    row_order INTEGER DEFAULT 0,
    UNIQUE(product_id, section, row_order)
);
CREATE INDEX IF NOT EXISTS idx_field_prod ON fields(product_id);
CREATE INDEX IF NOT EXISTS idx_field_key ON fields(field_key);
CREATE INDEX IF NOT EXISTS idx_field_sec ON fields(section);

CREATE TABLE IF NOT EXISTS components (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    product_id INTEGER NOT NULL REFERENCES products(id) ON DELETE CASCADE,
    comp_idx INTEGER DEFAULT 0,
    name TEXT DEFAULT '',
    cas TEXT DEFAULT '',
    conc TEXT DEFAULT '',
    raw_name TEXT DEFAULT '',
    raw_cas TEXT DEFAULT '',
    raw_conc TEXT DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_comp_prod ON components(product_id);

CREATE TABLE IF NOT EXISTS pictograms (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    product_id INTEGER NOT NULL REFERENCES products(id) ON DELETE CASCADE,
    section INTEGER DEFAULT 0,
    ext TEXT DEFAULT 'png',
    blob BLOB NOT NULL,
    width INTEGER DEFAULT 0,
    height INTEGER DEFAULT 0
);
CREATE INDEX IF NOT EXISTS idx_pic_prod ON pictograms(product_id);

CREATE TABLE IF NOT EXISTS field_dict (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    std_field TEXT NOT NULL,
    section INTEGER DEFAULT 0,
    display_order INTEGER DEFAULT 0,
    freq INTEGER DEFAULT 0,
    aliases_json TEXT DEFAULT '[]',
    status TEXT DEFAULT 'active',
    note TEXT DEFAULT '',
    UNIQUE(std_field, section)
);
CREATE INDEX IF NOT EXISTS idx_fdict_sec ON field_dict(section);

CREATE TABLE IF NOT EXISTS revisions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    product_id INTEGER REFERENCES products(id) ON DELETE CASCADE,
    action TEXT DEFAULT '',
    old_json TEXT DEFAULT '',
    created_at TEXT DEFAULT ''
);
"""

# 节标题标准名 (16 节 + 0 页眉页脚), 用于回读重建 full_title 兜底
SECTION_TITLES = {
    0: "页眉/页脚",
    1: "物料及供应商标识",
    2: "危险性概述",
    3: "成分/组成信息",
    4: "急救措施",
    5: "消防措施",
    6: "泄漏应急处理",
    7: "操作处置与储存",
    8: "接触控制和个体防护",
    9: "理化特性",
    10: "稳定性和反应性",
    11: "毒理学信息",
    12: "生态学信息",
    13: "废弃处置",
    14: "运输信息",
    15: "法规信息",
    16: "其他信息",
}


_COMP_RAW_MIGRATIONS = (
    ("ALTER TABLE components ADD COLUMN raw_name TEXT DEFAULT ''", "raw_name"),
    ("ALTER TABLE components ADD COLUMN raw_cas TEXT DEFAULT ''", "raw_cas"),
    ("ALTER TABLE components ADD COLUMN raw_conc TEXT DEFAULT ''", "raw_conc"),
)


def connect(db_path):
    """打开 (不存在则新建) 总库并建表."""
    import sqlite3
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    conn.executescript(SCHEMA)
    # 幂等迁移: 旧库 components 表补 raw_* 列 (已存在则跳过)
    existing = {r["name"] for r in conn.execute(
        "PRAGMA table_info(components)")}
    for sql, col in _COMP_RAW_MIGRATIONS:
        if col not in existing:
            try:
                conn.execute(sql)
            except sqlite3.OperationalError:
                pass  # 并发建表竞态, 忽略
    conn.commit()
    return conn
