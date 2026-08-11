# -*- coding: utf-8 -*-
"""MSDS 总库核心读写: 写入(收敛) / 回读(展开) / 物化宽表 / 检索 / 增删改.

写入 (导入 -> 全量入库):
  - products 一行 + fields 全行 (节标题/sub/field/note 逐行入库, 禁止缺失)
  - field_key = 字典收敛后的标准字段; raw_label = 原始表达 (原样)
  - components + pictograms blob 完整入库
回读 (型号 -> ParseResult 呈现):
  - 从 fields 重建 SectionData, raw_label 作原始标签, 与原文一致
物化宽表 (检索/对比/导出层):
  - 列 = 标准字段 field_key (按节内首现序), 行 = 型号, 值 = 聚合
"""
from __future__ import annotations

import json
import re
import sys
import time
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))
if str(_ROOT / "批量化读取") not in sys.path:
    sys.path.insert(0, str(_ROOT / "批量化读取"))

from build_db import _nk, normalize_field  # 复用静态归一 (P2 前兜底)

from .schema import SECTION_TITLES, connect

_LABEL_STRIP_RE = re.compile(r"^[·•\-—–\s]+")
_SEQ_STRIP_RE = re.compile(r"^\d+(\.\d+)*[\.\s]*")


def strip_label_prefix(raw):
    """去标签前缀 (项目符号/自动序号), 返回纯标签."""
    t = _LABEL_STRIP_RE.sub("", (raw or "").strip())
    t = _SEQ_STRIP_RE.sub("", t).strip()
    return t


def iter_field_dict(conn, section=None):
    """迭代 field_dict 全部记录 (可选按节过滤): (std_field, alias)."""
    if section is None:
        rows = conn.execute("SELECT std_field, section, aliases_json FROM field_dict")
    else:
        rows = conn.execute(
            "SELECT std_field, section, aliases_json FROM field_dict "
            "WHERE section=? OR section=0", (section,))
    for r in rows:
        try:
            aliases = json.loads(r["aliases_json"] or "[]")
        except Exception:
            aliases = []
        for a in aliases:
            yield r["std_field"], a


def resolve_field_key(conn, section, raw_label):
    """原始标签 -> 标准字段 key (字典优先, 静态词典兜底, 原样最后兜底)."""
    if not raw_label:
        return ""
    lab = strip_label_prefix(raw_label)
    nk = _nk(lab)
    if conn is not None:
        for std, alias in iter_field_dict(conn, section):
            if nk and nk == _nk(alias):
                return std
    return normalize_field(lab)


def _now():
    return time.strftime("%Y-%m-%d %H:%M:%S")


_MODEL_FROM_NAME_RE = re.compile(r"^([A-Za-z]{2,4}[-－]?\d+[A-Za-z]?[-－]?\w*)")


def extract_model_name(file_name):
    """从文件名提取型号: 'BL-8085 msds_CN 国彩.docx' -> 'BL-8085'."""
    stem = Path(file_name).stem.strip()
    m = _MODEL_FROM_NAME_RE.match(stem)
    if m:
        return m.group(1).replace("－", "-").strip()
    return stem


# ============================================================
# 写入: 全量入库 (收敛 field_key, 保留 raw_label)
# ============================================================

def add_product(conn, result, *, category="", model_name=None,
                file_name=None, note=""):
    """把一个 ParseResult 完整写入总库. 返回 (product_id, 状态).

    状态: 'added' 新增 | 'exists_sha' sha256重复 | 'exists_model' 型号重复.
    所有 fields/components/pictograms 全量写入, 不丢任何行.
    """
    model = (model_name or extract_model_name(result.file_name)).strip()
    cat_id = None
    if category:
        cur = conn.execute("SELECT id FROM categories WHERE name=?", (category,)).fetchone()
        if cur:
            cat_id = cur["id"]
        else:
            cur = conn.execute(
                "INSERT INTO categories (name, sort_order) VALUES (?, "
                "(SELECT COALESCE(MAX(sort_order),0)+1 FROM categories))", (category,))
            cat_id = cur.lastrowid

    if result.sha256:
        dup = conn.execute("SELECT id FROM products WHERE sha256=? AND active=1",
                           (result.sha256,)).fetchone()
        if dup:
            return dup["id"], "exists_sha"
    dup = conn.execute("SELECT id FROM products WHERE category_id IS ? AND model_name=? "
                       "AND active=1", (cat_id, model)).fetchone()
    if dup:
        return dup["id"], "exists_model"

    fn = file_name or result.file_name
    # 产品元数据: 从 S0 页眉字段提取 产品名/版本/修订日期 (页眉页脚全文在
    # S0 fields/note 中完整入库, 不丢失)
    prod_name = ver = rev = ""
    sec0 = result.sections.get(0)
    if sec0:
        for row in sec0.iter_rows():
            lab = (row.label or "").lower()
            if not row.value:
                continue
            if "产品名称" in lab or "product" in lab:
                prod_name = prod_name or row.value
            elif "version" in lab or "版本" in lab:
                ver = ver or row.value
            elif "修订" in lab or "revision" in lab or "prevision" in lab:
                rev = rev or row.value
    cur = conn.execute(
        "INSERT INTO products (category_id, model_name, file_name, file_path, sha256, "
        "language, template_vendor, product_name, version, revision_date, "
        "active, created_at, updated_at, note) VALUES (?,?,?,?,?,?,?,?,?,?,1,?,?,?)",
        (cat_id, model, fn, result.file_path, result.sha256,
         getattr(result, "language", ""), getattr(result, "template_vendor", ""),
         prod_name, ver, rev, _now(), _now(), note))
    pid = cur.lastrowid

    # 全量 fields 入库 (含 0 节页眉页脚, 含 section/sub/note 标题与说明)
    # row_order: 节内重编号 (section 标题行=0, 其余从 1 递增), 避免与
    # iter_rows 的 index 语义 (kind 隔离) 冲突 UNIQUE(product, section, row_order).
    for sec_num, sec in sorted(result.sections.items()):
        ro = 0
        for row in sec.iter_rows():
            if row.kind == "section":
                ro_row = 0
            else:
                ro += 1
                ro_row = ro
            key = resolve_field_key(conn, sec_num, row.label)
            conn.execute(
                "INSERT INTO fields (product_id, section, seq, field_key, raw_label, "
                "value, kind, editable, row_order) VALUES (?,?,?,?,?,?,?,?,?)",
                (pid, sec_num, row.seq, key, row.label, row.value,
                 row.kind, 1 if row.editable else 0, ro_row))

    sec3 = result.sections.get(3)
    if sec3:
        for i, c in enumerate(sec3.components, 1):
            conn.execute(
                "INSERT INTO components (product_id, comp_idx, name, cas, conc, "
                "raw_name, raw_cas, raw_conc) VALUES (?,?,?,?,?,?,?,?)",
                (pid, i, c.name, c.cas, c.conc,
                 c.raw_name, c.raw_cas, c.raw_conc))
        # 成分表头持久化 (禁止缺失): component_header 作为独立行入库.
        # kind='component_header' 不参与 field 语义 (field_key 空), 供
        # 检索/回读/导出; row_order 接 S3 末尾 (UNIQUE product+section+order).
        if sec3.component_header:
            mx = conn.execute(
                "SELECT COALESCE(MAX(row_order),0) m FROM fields "
                "WHERE product_id=? AND section=3", (pid,)).fetchone()["m"]
            conn.execute(
                "INSERT INTO fields (product_id, section, seq, field_key, "
                "raw_label, value, kind, editable, row_order) "
                "VALUES (?,3,'','',?,?, 'component_header', 0, ?)",
                (pid, sec3.component_header, sec3.component_header, mx + 1))

    for im in result.images:
        conn.execute(
            "INSERT INTO pictograms (product_id, section, ext, blob, width, height) "
            "VALUES (?,?,?,?,?,?)",
            (pid, im.section, im.ext, im.blob, im.width, im.height))

    conn.execute("INSERT INTO revisions (product_id, action, created_at) "
                 "VALUES (?,?,?)", (pid, "add", _now()))
    conn.commit()
    return pid, "added"


def update_product_meta(conn, product_id, *, category=None,
                        model_name=None, note=None):
    """更新型号归属/名称/备注 (不重写内容)."""
    if category is not None:
        cur = conn.execute("SELECT id FROM categories WHERE name=?", (category,)).fetchone()
        cat_id = cur["id"] if cur else conn.execute(
            "INSERT INTO categories (name, sort_order) VALUES (?, "
            "(SELECT COALESCE(MAX(sort_order),0)+1 FROM categories))",
            (category,)).lastrowid
        conn.execute("UPDATE products SET category_id=? WHERE id=?", (cat_id, product_id))
    if model_name is not None:
        conn.execute("UPDATE products SET model_name=? WHERE id=?", (model_name, product_id))
    if note is not None:
        conn.execute("UPDATE products SET note=? WHERE id=?", (note, product_id))
    conn.execute("UPDATE products SET updated_at=? WHERE id=?", (_now(), product_id))
    conn.commit()


def delete_product(conn, product_id, hard=False):
    """删除型号. 默认软删除 (active=0); hard=True 物理删 (级联清子表)."""
    if hard:
        snap = export_product_json(conn, product_id)
        conn.execute("INSERT INTO revisions (product_id, action, old_json, created_at) "
                     "VALUES (?,?,?,?)", (product_id, "delete", snap, _now()))
        conn.execute("DELETE FROM products WHERE id=?", (product_id,))
    else:
        conn.execute("UPDATE products SET active=0, updated_at=? WHERE id=?",
                     (_now(), product_id))
        conn.execute("INSERT INTO revisions (product_id, action, created_at) "
                     "VALUES (?,?,?,?)", (product_id, "delete", _now()))
    conn.commit()


# ============================================================
# 查询: 回读 (展开 raw_label) / 物化宽表 / 检索
# ============================================================

def list_categories(conn):
    return [dict(r) for r in conn.execute(
        "SELECT id, name, sort_order FROM categories ORDER BY sort_order, name")]


def list_products(conn, category_id=None, active_only=True, unassigned=False):
    """型号列表. category_id=None 默认不限类目; unassigned=True 仅取未分类."""
    sql = ("SELECT p.*, c.name AS category_name "
           "FROM products p LEFT JOIN categories c ON p.category_id=c.id WHERE 1=1")
    args = []
    if unassigned:
        sql += " AND p.category_id IS NULL"
    elif category_id is not None:
        sql += " AND p.category_id=?"
        args.append(category_id)
    if active_only:
        sql += " AND p.active=1"
    sql += " ORDER BY p.model_name"
    return [dict(r) for r in conn.execute(sql, args)]


def get_product(conn, product_id):
    r = conn.execute(
        "SELECT p.*, c.name AS category_name FROM products p "
        "LEFT JOIN categories c ON p.category_id=c.id WHERE p.id=?", (product_id,)).fetchone()
    return dict(r) if r else None


def get_fields(conn, product_id):
    return [dict(r) for r in conn.execute(
        "SELECT * FROM fields WHERE product_id=? ORDER BY section, row_order",
        (product_id,))]


def get_components(conn, product_id):
    return [dict(r) for r in conn.execute(
        "SELECT * FROM components WHERE product_id=? ORDER BY comp_idx", (product_id,))]


def get_pictograms(conn, product_id):
    return [dict(r) for r in conn.execute(
        "SELECT * FROM pictograms WHERE product_id=? ORDER BY section, id", (product_id,))]


def get_field_keys(conn, product_id):
    """回读用: (section, field_key) -> 首个 raw_label (原始表达)."""
    out = {}
    for r in conn.execute(
            "SELECT section, field_key, raw_label FROM fields "
            "WHERE product_id=? AND field_key!='' ORDER BY row_order", (product_id,)):
        out.setdefault((r["section"], r["field_key"]), r["raw_label"])
    return out


def export_product_json(conn, product_id):
    """型号完整快照 (products + fields + components + pictograms 摘要), 供版本历史."""
    p = get_product(conn, product_id)
    if not p:
        return "{}"
    data = {
        "product": p,
        "fields": get_fields(conn, product_id),
        "components": get_components(conn, product_id),
        "pictograms": [
            {"section": x["section"], "ext": x["ext"], "size": len(x["blob"])}
            for x in get_pictograms(conn, product_id)],
    }
    return json.dumps(data, ensure_ascii=False, default=str)


# ---- 物化宽表 (检索/对比/导出层) ----

def materialize_wide(conn, active_only=True):
    """物化宽表: 行=型号, 列=标准字段 + 成分列.

    返回: {'columns': [列头], 'col_keys': [列key], 'rows': [(型号, {列key: 值})],
           'comp_max': 成分最大数}
    列序 = (节号, 节内首现 row_order); 值聚合: 同名多行 换行 合并 + 行级去重.
    """
    cols = []
    seen_col = {}
    pos = {}
    rows = list_products(conn, active_only=active_only)

    for p in rows:
        for r in conn.execute(
                "SELECT section, field_key, row_order FROM fields "
                "WHERE product_id=? AND kind='field' AND field_key!='' "
                "AND section!=0 ORDER BY section, row_order", (p["id"],)):
            k = (r["section"], r["field_key"])
            if k not in seen_col:
                seen_col[k] = len(cols)
                cols.append({"section": r["section"], "field_key": r["field_key"]})
            pos[k] = min(pos.get(k, 10 ** 9), r["row_order"])

    cols.sort(key=lambda c: (c["section"], pos[(c["section"], c["field_key"])]))
    columns = [f"{c['section']}.{c['field_key']}" if c["section"] else c["field_key"]
               for c in cols]
    col_keys = [(c["section"], c["field_key"]) for c in cols]

    comp_max = 0
    for p in rows:
        n = conn.execute("SELECT COUNT(*) c FROM components WHERE product_id=?",
                         (p["id"],)).fetchone()["c"]
        comp_max = max(comp_max, n)

    out_rows = []
    for p in rows:
        cell = {}
        for r in conn.execute(
                "SELECT section, field_key, value FROM fields "
                "WHERE product_id=? AND kind='field' AND field_key!='' "
                "AND section!=0 ORDER BY section, row_order", (p["id"],)):
            k = (r["section"], r["field_key"])
            v = (r["value"] or "").strip()
            if not v:
                continue
            if k in cell:
                if v not in cell[k].split("\n"):
                    cell[k] = (cell[k] + "\n" + v).strip()
            else:
                cell[k] = v
        for i, c in enumerate(get_components(conn, p["id"]), 1):
            for part, val in (("名称", c["name"]), ("CAS", c["cas"]), ("含量", c["conc"])):
                if val:
                    cell[("comp", i, part)] = val
        out_rows.append((p["id"], p["model_name"], cell))

    return {"columns": columns, "col_keys": col_keys, "rows": out_rows,
            "comp_max": comp_max}


def search_products(conn, query, *, scope="all"):
    """总库检索. scope: all | label | value | section | model.

    命中层双轨: field_key(标准) 与 raw_label(原始) 都参与匹配 ->
    搜 '毒性' 既能命中标准字段 生态毒性, 也能命中原始表达 12.1毒性.
    """
    q = _nk(query)
    if not q:
        return []
    pid_list = set()
    for p in list_products(conn):
        if scope == "model":
            if q in _nk(p["model_name"]):
                pid_list.add(p["id"])
            continue
        if scope == "section":
            m = re.search(r"\d+", query)
            sec = int(m.group(0)) if m else None
            if sec is not None and any(f["section"] == sec for f in get_fields(conn, p["id"])):
                pid_list.add(p["id"])
            continue
        # 成分检索 (名称/CAS/含量)
        for c in get_components(conn, p["id"]):
            if q in _nk(f"{c['name']} {c['cas']} {c['conc']}"):
                pid_list.add(p["id"])
                break
        if pid_list and p["id"] in pid_list:
            continue
        for f in get_fields(conn, p["id"]):
            if scope == "label":
                hay = f"{_nk(f['raw_label'])} {_nk(f['field_key'])}"
            elif scope == "value":
                hay = _nk(f["value"])
            else:
                hay = (f"{_nk(f['raw_label'])} {_nk(f['field_key'])} {_nk(f['value'])} "
                       f"{_nk(p['model_name'])} {_nk(p['file_name'])}")
            if q in hay:
                pid_list.add(p["id"])
                break
    out = []
    for pid in pid_list:
        p = get_product(conn, pid)
        if p:
            out.append(p)
    out.sort(key=lambda x: (x.get("category_name") or "", x["model_name"]))
    return out


# ---- 统计 ----

def db_stats(conn):
    return {
        "categories": conn.execute("SELECT COUNT(*) c FROM categories").fetchone()["c"],
        "products": conn.execute("SELECT COUNT(*) c FROM products WHERE active=1").fetchone()["c"],
        "products_all": conn.execute("SELECT COUNT(*) c FROM products").fetchone()["c"],
        "fields": conn.execute("SELECT COUNT(*) c FROM fields").fetchone()["c"],
        "components": conn.execute("SELECT COUNT(*) c FROM components").fetchone()["c"],
        "pictograms": conn.execute("SELECT COUNT(*) c FROM pictograms").fetchone()["c"],
        "field_dict": conn.execute("SELECT COUNT(*) c FROM field_dict").fetchone()["c"],
    }


# ============================================================
# 回读: fields → ParseResult (raw_label 原始表达, 走 reader 标准结构)
# ============================================================

def rebuild_product(conn, product_id):
    """从总库重建 ParseResult (供 GUI 以 reader 标准结构呈现/编辑).

    重建规则 (保真, 不重拆):
      - 所有非 section 行 → FieldData (label=seq+raw_label 拼回原始形态,
        sub 行 value 空 editable=False, note 行 label 空)
      - 字段顺序按 row_order (order 全 field → iter_rows 直出, 零重拆)
      - components / pictograms 原样重建
      - raw_label = 原始表达 (读取时用原有检索字段)
    """
    from core.structure import (FieldData, ComponentData, ImageData,
                                ParseResult, SectionData)

    p = get_product(conn, product_id)
    if not p:
        return None

    sec_groups: dict[int, list] = {}
    for f in get_fields(conn, product_id):
        sec_groups.setdefault(f["section"], []).append(f)

    sections: dict[int, SectionData] = {}
    for n, rows in sorted(sec_groups.items()):
        sec = SectionData(number=n, title=SECTION_TITLES.get(n, ""),
                          full_title=f"{n}.{SECTION_TITLES.get(n, '')}" if n else SECTION_TITLES.get(n, ""))
        fields: list[FieldData] = []
        for f in sorted(rows, key=lambda x: x["row_order"]):
            kind = f["kind"]
            if kind == "component_header":
                # 成分表头回读 (禁止缺失): 恢复 sec.component_header
                sec.component_header = f["value"] or ""
                continue
            if kind == "section":
                sec.title = f["raw_label"]
                sec.full_title = f["raw_label"]
                continue
            lab = f"{f['seq']} {f['raw_label']}".strip() if f["seq"] else f["raw_label"]
            if kind == "sub":
                fields.append(FieldData(label=lab, value="", editable=False))
            elif kind == "note":
                fields.append(FieldData(label="", value=f["value"],
                                        editable=bool(f["editable"])))
            else:  # field
                fields.append(FieldData(label=lab, value=f["value"],
                                        editable=bool(f["editable"])))
            sec.order.append("field")
        sec.fields = fields
        if n == 3:
            comps = get_components(conn, product_id)
            sec.components = [ComponentData(name=c["name"], cas=c["cas"],
                                            conc=c["conc"],
                                            raw_name=c["raw_name"] or "",
                                            raw_cas=c["raw_cas"] or "",
                                            raw_conc=c["raw_conc"] or "")
                              for c in comps]
            sec.is_component_table = bool(comps)
        sections[n] = sec

    images = [ImageData(blob=x["blob"], ext=x["ext"], section=x["section"],
                        width=x["width"], height=x["height"])
              for x in get_pictograms(conn, product_id)]

    return ParseResult(
        file_path=p.get("file_path", ""), file_name=p.get("file_name", ""),
        sha256=p.get("sha256", ""), sections=sections, images=images,
        sections_count=len(sections))


def update_field_value(conn, field_id: int, value: str) -> None:
    """更新某字段行的值 (GUI 编辑保存). 同步刷新型号 updated_at."""
    conn.execute("UPDATE fields SET value=? WHERE id=?", (value, field_id))
    conn.execute("UPDATE products SET updated_at=? WHERE id=(SELECT product_id FROM fields WHERE id=?)",
                 (_now(), field_id))
    conn.commit()
