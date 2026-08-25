#!/usr/bin/env python3
"""按 CAS 查询 REACH SVHC 数据库，并列出无 CAS 待审查物质。"""

from __future__ import annotations

import argparse
import re
import sqlite3
import sys
from pathlib import Path


DEFAULT_DB = Path(
    r"F:\正式项目与模块化内容\冠志\MSDS\04-推断引擎\正式法律法规库\物质限制清单\REACH-SVHC-253项-物质数据库-2026-02-04.db"
)
CAS_RE = re.compile(r"^\d{2,7}-\d{2,7}-\d$")


def query_database(db_path: Path, cas: str) -> tuple[list[tuple], list[tuple]]:
    if not db_path.exists():
        raise FileNotFoundError(f"数据库不存在：{db_path}")

    con = sqlite3.connect(f"file:{db_path.as_posix()}?mode=ro", uri=True)
    try:
        main_rows = con.execute(
            """
            SELECT "中文名称", "英文名称", "物质描述", "EC号", "CAS号"
            FROM reach_svhc_main
            WHERE instr(',' || replace(replace("CAS号", char(10), ','), ' ', '') || ',',
                        ',' || ? || ',') > 0
               OR trim("CAS号") = ?
            ORDER BY rowid
            """,
            (cas, cas),
        ).fetchall()
        no_cas_rows = con.execute(
            """
            SELECT "中文名称", "英文名称"
            FROM reach_svhc_missing_identifier
            WHERE trim(coalesce("CAS号", '')) = ''
            ORDER BY rowid
            """
        ).fetchall()
        return main_rows, no_cas_rows
    finally:
        con.close()


def print_result(cas: str, main_rows: list[tuple], no_cas_rows: list[tuple]) -> None:
    print(f"CAS查询：{cas}")
    print()

    if main_rows:
        print("警告：该 CAS 对应 REACH SVHC 候选清单物质。")
        for index, row in enumerate(main_rows, 1):
            cn, en, description, ec, db_cas = row
            print(f"命中记录 {index}：")
            print(f"中文名称：{cn or '（空）'}")
            print(f"英文名称：{en or '（空）'}")
            print(f"物质描述：{description or '（空）'}")
            print(f"EC号：{ec or '（空）'}")
            print(f"CAS号：{db_cas or '（空）'}")
            if index != len(main_rows):
                print()
    else:
        print("结果：该 CAS 未在 REACH SVHC 有 CAS 主库中找到。")

    print()
    print("无CAS待审查物质")
    if no_cas_rows:
        for cn, en in no_cas_rows:
            print(f"中文名称：{cn or '（空）'}；英文名称：{en or '（空）'}")
    else:
        print("（无）")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="查询 REACH SVHC SQLite 数据库，并列出无 CAS 待审查物质。"
    )
    parser.add_argument("cas", nargs="?", help="要查询的 CAS 号，例如 110-71-4")
    parser.add_argument("--db", type=Path, default=DEFAULT_DB, help="SQLite 数据库路径")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    cas = (args.cas or input("请输入CAS号：")).strip()
    if not CAS_RE.fullmatch(cas):
        print(f"输入的 CAS 号格式不正确：{cas}", file=sys.stderr)
        return 2
    try:
        main_rows, no_cas_rows = query_database(args.db, cas)
    except (FileNotFoundError, sqlite3.Error) as exc:
        print(f"数据库查询失败：{exc}", file=sys.stderr)
        return 1
    print_result(cas, main_rows, no_cas_rows)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
