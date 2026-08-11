# -*- coding: utf-8 -*-
"""正式筛选库构建器 — 从 MSDS 原始数据源构建 统一库/宽表库/差异库."""
from __future__ import annotations
import json, re, sqlite3, sys, time
from pathlib import Path
from collections import defaultdict

_ROOT = Path(__file__).resolve().parent.parent
for p in (_ROOT, _ROOT / "批量化读取"):
    if str(p) not in sys.path:
        sys.path.insert(0, str(p))
from build_db import normalize_field

DEFAULT_OUT = Path(r"F:\正式项目与模块化内容\Word 覆写模块\数据库\正式筛选库")

STD_COLUMNS_ZH = [
    (1, "产品名称"), (1, "中文名称"), (1, "化学品分类"), (1, "供应商信息"),
    (1, "供应商名称"), (1, "供应商地址"), (1, "电话"), (1, "产品使用建议"), (1, "传真"),
    (2, "物质或混合物分类"), (2, "GHS危险性类别"), (2, "防范说明"), (2, "危害性说明"),
    (3, "产品类型"),
    (4, "一般措施"), (4, "误服"), (4, "接触眼睛"), (4, "接触皮肤"), (4, "吸入"),
    (5, "合适的灭火剂"), (5, "不合适的灭火剂"), (5, "物质或混合物的特殊危害"), (5, "消防预防措施和保护设备"),
    (6, "个人预防措施、应急程序"), (6, "环境保护措施"),
    (7, "安全操作防范"), (7, "安全储存条件"),
    (8, "眼睛防护"), (8, "身体防护"), (8, "呼吸系统防护"), (8, "手部防护"),
    (8, "防护手套的合适材料"), (8, "建议"), (8, "暴露限值"),
    (9, "其他信息"), (9, "外观"), (9, "水溶性"), (9, "密度"), (9, "相对密度"),
    (9, "动力粘度"), (9, "蒸发速率"), (9, "表面张力"), (9, "初沸点"), (9, "相对蒸气密度"),
    (9, "辛醇/水分配系数"), (9, "燃烧值"), (9, "离子性"), (9, "固体含量"), (9, "闪点"),
    (9, "可燃性"), (9, "引燃温度"), (9, "嗅觉阀值"), (9, "自燃温度"), (9, "分解温度"),
    (9, "粉尘爆炸级别"), (9, "爆炸特性"),
    (10, "化学稳定性"), (10, "可能的危害反应"), (10, "危险分解产物"),
    (11, "急性毒性"), (11, "致敏性"), (11, "致癌性"), (11, "致突变性"),
    (11, "主要皮肤刺激性"), (11, "主要粘膜刺激性"), (11, "类似产品的风险评估数据"),
    (12, "生态毒性"), (12, "持久性和降解性"), (12, "其他"),
    (13, "处理方法"),
    (14, "公路和铁路运输"), (14, "海上运输"), (14, "空运"), (14, "用户特殊注意事项"),
    (15, "符合下列法规要求"), (15, "其它的规定"),
    (16, "其他信息"),
]

STD_COLUMNS_EN = [
    (1, "产品名称"), (1, "产品使用建议"), (1, "供应商名称"), (1, "供应商地址"), (1, "电话"), (1, "传真"),
    (1, "供应商信息"), (1, "化学品分类"),
    (3, "产品类型"),
    (4, "一般措施"), (4, "吸入"), (4, "接触眼睛"), (4, "误服"), (4, "接触皮肤"),
    (5, "合适的灭火剂"), (5, "消防预防措施和保护设备"),
    (6, "环境保护措施"), (6, "泄漏程序"),
    (7, "安全储存条件"), (7, "安全操作防范"),
    (8, "呼吸系统防护"), (8, "眼睛防护"), (8, "手部防护"), (8, "建议"), (8, "身体防护"),
    (8, "防护手套的合适材料"), (8, "暴露限值"),
    (9, "外观"), (9, "可燃性"), (9, "引燃温度"), (9, "自燃温度"), (9, "其他信息"),
    (9, "水溶性"), (9, "闪点"), (9, "固体含量"), (9, "分解温度"), (9, "爆炸特性"),
    (9, "粘度"), (9, "离子性"), (9, "pH值"), (9, "蒸发速率"), (9, "相对密度"),
    (9, "动力粘度"), (9, "初沸点"), (9, "辛醇/水分配系数"), (9, "粉尘爆炸级别"),
    (10, "危险分解产物"), (10, "化学稳定性"),
    (11, "致敏性"), (11, "急性毒性"), (11, "致突变性"), (11, "主要粘膜刺激性"),
    (11, "主要皮肤刺激性"), (11, "类似产品的风险评估数据"),
    (12, "生态毒性"), (12, "持久性和降解性"),
    (13, "处理方法"), (13, "空容器注意事项"),
    (14, "公路和铁路运输"), (14, "海上运输"), (14, "空运"), (14, "环境危害"),
    (14, "联合国编号"),
    (15, "其它的规定"), (15, "符合下列法规要求"),
    (16, "其他信息"),
]

DIFF_FIELDS = ["供应商信息", "供应商名称", "供应商地址", "电话", "传真", "产品使用建议", "产品名称"]


def _merge_values(vals):
    cleaned = [v.strip() for v in vals if v and v.strip()]
    if not cleaned:
        return ""
    seen, uniq = set(), []
    for v in cleaned:
        if v not in seen:
            seen.add(v); uniq.append(v)
    return max(uniq, key=len) if uniq else ""


def _model_key(model):
    return (model or "").upper().replace(" ", "")


def _col_defs(cols):
    out = []
    for sec, label in cols:
        out.append(chr(34) + str(sec) + '_' + label + chr(34) + ' TEXT DEFAULT ' + chr(39)*2)
    return out


def build_screened_db(parsed_pkl, out_dir):
    import pickle
    records = pickle.load(open(parsed_pkl, "rb"))
    out_dir.mkdir(parents=True, exist_ok=True)
    db_path = out_dir / "正式筛选库.db"
    if db_path.exists():
        db_path.unlink()

    zh_cols = _col_defs(STD_COLUMNS_ZH)
    en_cols = _col_defs(STD_COLUMNS_EN)
    diff_cols = [chr(34) + f + chr(34) + " TEXT DEFAULT " + chr(39)*2 for f in DIFF_FIELDS]

    def col_block(cols):
        return ",\n    ".join(cols)

    schema = f"""
CREATE TABLE IF NOT EXISTS zh_products (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    model TEXT NOT NULL UNIQUE,
    vendors TEXT DEFAULT '',
    files TEXT DEFAULT '',
    {col_block(zh_cols)},
    created_at TEXT DEFAULT ''
);
CREATE TABLE IF NOT EXISTS en_products (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    model TEXT NOT NULL UNIQUE,
    vendors TEXT DEFAULT '',
    files TEXT DEFAULT '',
    {col_block(en_cols)},
    created_at TEXT DEFAULT ''
);
CREATE TABLE IF NOT EXISTS wide_rows (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    model TEXT NOT NULL,
    lang TEXT NOT NULL,
    section INTEGER,
    std_label TEXT DEFAULT '',
    raw_label TEXT DEFAULT '',
    value TEXT DEFAULT '',
    vendor TEXT DEFAULT '',
    seq TEXT DEFAULT '',
    kind TEXT DEFAULT 'field'
);
CREATE INDEX IF NOT EXISTS idx_wide_model ON wide_rows(model);
CREATE INDEX IF NOT EXISTS idx_wide_std ON wide_rows(std_label);
CREATE INDEX IF NOT EXISTS idx_wide_raw ON wide_rows(raw_label);
CREATE TABLE IF NOT EXISTS diff_products (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    model TEXT NOT NULL,
    lang TEXT NOT NULL,
    vendor TEXT NOT NULL,
    {col_block(diff_cols)},
    created_at TEXT DEFAULT '',
    UNIQUE(model, lang, vendor)
);
"""
    conn = sqlite3.connect(str(db_path))
    conn.executescript(schema)

    by_lang_model = defaultdict(list)
    for r in records:
        by_lang_model[(r["lang"], _model_key(r["model"]))].append(r)

    stats = {"zh_models": 0, "en_models": 0, "wide": 0, "diff": 0}

    for (lang, model_key), recs in sorted(by_lang_model.items()):
        col_list = STD_COLUMNS_ZH if lang == "zh" else STD_COLUMNS_EN
        table = "zh_products" if lang == "zh" else "en_products"

        merged = defaultdict(list)
        vendors = set()
        files = []
        for r in recs:
            vendors.add(r["vendor"])
            files.append(r["file"])
            for f in r["fields"]:
                if f["kind"] not in ("field", "note"):
                    continue
                std = normalize_field(f["label"])
                val = (f["value"] or "").strip()
                if val:
                    conn.execute(
                        "INSERT INTO wide_rows (model, lang, section, std_label, raw_label, value, vendor, seq, kind) "
                        "VALUES (?,?,?,?,?,?,?,?,?)",
                        (r["model"], lang, f["sec"], std, f["label"], val,
                         r["vendor"], f["seq"], f["kind"]))
                    stats["wide"] += 1
                if std:
                    merged[(f["sec"], std)].append(val)

        row_vals = {}
        for sec, label in col_list:
            key = (sec, label)
            row_vals[f"{sec}_{label}"] = _merge_values(merged.get(key, []))
        row_vals["model"] = model_key
        row_vals["vendors"] = ",".join(sorted(vendors))
        row_vals["files"] = json.dumps(files, ensure_ascii=False)
        row_vals["created_at"] = time.strftime("%Y-%m-%d %H:%M:%S")

        cols = list(row_vals.keys())
        qcols = [f'"{c}"' for c in cols]
        ph = ",".join("?" for _ in cols)
        conn.execute(f"INSERT INTO {table} ({','.join(qcols)}) VALUES ({ph})",
                     [row_vals[c] for c in cols])
        stats["zh_models" if lang == "zh" else "en_models"] += 1

        by_vendor = defaultdict(list)
        for r in recs:
            by_vendor[r["vendor"]].append(r)
        for vendor, vrecs in by_vendor.items():
            dvals = defaultdict(list)
            for r in vrecs:
                for f in r["fields"]:
                    std = normalize_field(f["label"])
                    if std in DIFF_FIELDS:
                        dvals[std].append((f["value"] or "").strip())
            drow = {"model": model_key, "lang": lang, "vendor": vendor,
                    "created_at": time.strftime("%Y-%m-%d %H:%M:%S")}
            for f in DIFF_FIELDS:
                drow[f] = _merge_values(dvals.get(f, []))
            cols = list(drow.keys())
            qcols = [f'"{c}"' for c in cols]
            ph = ",".join("?" for _ in cols)
            conn.execute(f"INSERT INTO diff_products ({','.join(qcols)}) VALUES ({ph})",
                         [drow[c] for c in cols])
            stats["diff"] += 1

    conn.commit()
    conn.close()
    stats["out"] = str(db_path)
    return stats


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    argv = sys.argv[1:]
    parsed = argv[0] if argv else None
    out = DEFAULT_OUT
    if "-o" in argv:
        i = argv.index("-o")
        if i + 1 < len(argv):
            out = Path(argv[i + 1])
    if not parsed:
        print("用法: python build_screened_db.py <parsed.pkl> [-o 输出目录]")
        sys.exit(2)
    t0 = time.time()
    stats = build_screened_db(Path(parsed), out)
    print(f"建库完成 ({time.time()-t0:.0f}s): {stats}")
