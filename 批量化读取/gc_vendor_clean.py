# -*- coding: utf-8 -*-
"""国彩供应商清洗 — 统一库内将国彩来源混入的旧冠志供应商值替换为国彩标准值.

规则: 只动统一库 (zh_products / en_products / diff_products), 宽表库 wide_rows 保留原始表述不动.
替换依据: 国彩(Guocai)品牌的标准供应商信息; 任何含冠志标记的值都视为污染.
"""
from __future__ import annotations
import sqlite3
import sys
from pathlib import Path

DEFAULT_DB = Path(r"F:\正式项目与模块化内容\Word 覆写模块\数据库\正式筛选库\正式筛选库.db")

# 国彩标准供应商信息
GC_STD_ZH = {
    "供应商名称": "英德市国彩精细化工有限公司",
    "供应商地址": "广东省英德市白沙镇太平村更古坑凯迪工业园区",
    "电话": "86-763-2811205",
    "传真": "86-763-2811024",
}
GC_STD_EN = {
    "供应商名称": "YINGDE GUOCAI FINE CHEMICAL CO., LTD.",
    "供应商地址": "广东省英德市白沙镇太平村更古坑凯迪工业园区",
    "电话": "86-763-2811205",
    "传真": "86-763-2811024",
}
# 冠志标记: 命中任一即视为冠志值
GZ_MARKS = ["冠志", "86-20-82567990", "86-20-32214789", "萝岗区科学城", "科学城掬泉路"]
DIFF_FIELDS = ["供应商名称", "供应商地址", "电话", "传真"]


def _is_gz_value(val: str) -> bool:
    return bool(val) and any(m in val for m in GZ_MARKS)


def _std_for(lang: str) -> dict:
    return GC_STD_ZH if lang == "zh" else GC_STD_EN


def clean_diff_products(conn: sqlite3.Connection) -> int:
    """diff_products 国彩行的污染替换. 返回替换字段数."""
    n = 0
    rows = conn.execute(
        "SELECT model, lang FROM diff_products WHERE vendor='国彩'"
    ).fetchall()
    for model, lang in rows:
        std = _std_for(lang)
        sets, params = [], []
        for f in DIFF_FIELDS:
            cur = conn.execute(
                f'SELECT "{f}" FROM diff_products WHERE model=? AND lang=? AND vendor=\'国彩\'',
                (model, lang)).fetchone()[0] or ""
            if _is_gz_value(cur):
                sets.append(f'"{f}"=?')
                params.append(std[f])
                n += 1
        if sets:
            params += [model, lang]
            conn.execute(
                f"UPDATE diff_products SET {', '.join(sets)} "
                f"WHERE model=? AND lang=? AND vendor='国彩'", params)
    return n


def clean_unified(conn: sqlite3.Connection, table: str, lang: str) -> int:
    """zh_products / en_products 统一行: 对 vendors 只含国彩(无冠志)的型号,
    供应商字段必须是国彩标准值 (防止源文件混入的冠志值被取为合并值)."""
    n = 0
    std = _std_for(lang)
    rows = conn.execute(
        f'SELECT model, vendors, "{DIFF_FIELDS[0]}" FROM {table}'
    ).fetchall()
    for model, vendors, _ in rows:
        vlist = (vendors or "").split(",")
        if "国彩" not in vlist or "冠志" in vlist:
            continue  # 双品牌或有冠志来源: 保持取更完整结果
        sets, params = [], []
        for f in DIFF_FIELDS:
            cur = conn.execute(f'SELECT "{f}" FROM {table} WHERE model=?', (model,)).fetchone()[0] or ""
            if _is_gz_value(cur):
                sets.append(f'"{f}"=?')
                params.append(std[f])
                n += 1
        if sets:
            params.append(model)
            conn.execute(f"UPDATE {table} SET {', '.join(sets)} WHERE model=?", params)
    return n


def clean_db(db_path: Path | str = DEFAULT_DB) -> dict:
    conn = sqlite3.connect(str(db_path))
    try:
        d_diff = clean_diff_products(conn)
        d_zh = clean_unified(conn, "zh_products", "zh")
        d_en = clean_unified(conn, "en_products", "en")
        conn.commit()
        return {"diff_products": d_diff, "zh_products": d_zh, "en_products": d_en}
    finally:
        conn.close()


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    db = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_DB
    print(f"清洗完成: {clean_db(db)}  (宽表库未动)")
