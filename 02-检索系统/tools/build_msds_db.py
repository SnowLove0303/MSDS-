# -*- coding: utf-8 -*-
"""MSDS SQLite 标准字段库 CLI (build_msds_db).

以「父子级 + 标签」结构表为模型, 把 MSDS docx / 透视总表 xlsx 入库为
SQLite 四表: schema_field 标准字段字典 / msds_model 型号主表 /
msds_field 明细长表 / msds_wide Schema 宽表.

用法:
  python tools/build_msds_db.py <db路径> <docx或目录...> [--from-xlsx 透视表.xlsx...]
  python tools/build_msds_db.py <db路径> --list                 # 列出全部型号
  python tools/build_msds_db.py <db路径> --model <型号> [--json|--tsv] [--sections 1,9]   # 按型号检索 (三级树)
  python tools/build_msds_db.py <db路径> --model-search <关键词> [--sections 1,9]          # 关键词检索库内数据
  python tools/build_msds_db.py <db路径> --query 关键词          # 检索
  python tools/build_msds_db.py <db路径> --wide <型号>           # 宽表一行
  python tools/build_msds_db.py <db路径> --init                 # 仅重建表结构

示例:
  python tools/build_msds_db.py "../../数据库/正式库/Data Base/msds_standard.db" ^
      "templates/MSDS_CN 国彩 模板.docx" --from-xlsx "../../数据库/正式库/标准字段数据库Excel/导出表/PEA-4139 MSDS_CN 冠志 模板_信息_20260817_085127.xlsx"
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from core.msds_db import (find_models, insert_docx, insert_pivot_xlsx,
                          list_models, model_detail, model_search, open_db,
                          render_model_json, render_model_tree,
                          render_model_tsv, search, wide_row,
                          refresh_s9_mapping, _model_of)
from core.docx_reader import read_msds


def _collect_docx(inputs: list[str]) -> list[Path]:
    """递归展开 docx 目录 (排除临时文件和生成产物)."""
    out: list[Path] = []
    for raw in inputs:
        p = Path(raw)
        if p.is_dir():
            out.extend(sorted(f for f in p.rglob("*.docx")
                              if not f.name.startswith("~$")
                              and "模板覆写输出" not in f.stem
                              and "标准化输出" not in f.stem
                              and "底版" not in f.stem))
        elif p.exists() and p.suffix.lower() == ".docx":
            if any(x in p.stem for x in ("模板覆写输出", "标准化输出", "底版")):
                print(f"  ⚠️ 跳过生成产物: {p}")
            else:
                out.append(p)
        else:
            print(f"  ⚠️ 跳过 (非 docx 或不存在): {p}")
    return sorted(set(out), key=lambda x: str(x).lower())


def _canonical_docx_files(files: list[Path]) -> list[Path]:
    """按型号去重，保证一型号只进入一条数据库记录。

    来源目录可能同时存在冠志/国彩版本、(2) 副本或重复下载文件。
    选择规则固定为：冠志来源优先、非副本优先、路径字典序兜底；
    这样同一批输入每次重建都得到同一份中文型号库。
    """
    chosen: dict[str, tuple[tuple[int, int, str], Path]] = {}
    skipped = 0
    for path in files:
        try:
            result = read_msds(path)
            model = _model_of(result, path.name).strip()
        except Exception as exc:
            print(f"  ⚠️ 型号识别失败，保留入库由后续报错处理: {path.name}: {exc}")
            continue
        if not model:
            continue
        text = str(path).lower()
        priority = (
            0 if ("冠志" in text or "guanzhi" in text) else 1,
            1 if "(2)" in path.stem else 0,
            text,
        )
        old = chosen.get(model)
        if old is None or priority < old[0]:
            if old is not None:
                skipped += 1
            chosen[model] = (priority, path)
        else:
            skipped += 1
    selected = sorted((item[1] for item in chosen.values()),
                      key=lambda p: str(p).lower())
    if skipped:
        print(f"  ↳ 同型号源文件去重: 保留 {len(selected)} 份，跳过 {skipped} 份")
    return selected


def _parse_sections(rest: list[str]) -> set[int] | None:
    """解析 --sections 1,3,9 → {1,3,9}; 缺省 None (全部)."""
    if "--sections" in rest:
        i = rest.index("--sections")
        if i + 1 < len(rest):
            s = {int(x) for x in rest[i + 1].split(",") if x.strip().isdigit()}
            return s or None
    return None


def main() -> int:
    args = sys.argv[1:]
    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        return 0
    db_path = args[0]
    rest = args[1:]

    # 结构写死: 每次运行强制校验冻结结构 (防其他线程/Agent 改动 schema/骨架)
    from core.msds_db import validate_structure
    try:
        validate_structure()
    except RuntimeError as exc:
        print(str(exc))
        return 3

    # 只读子命令
    db = open_db(db_path)
    if "--list" in rest or "--models" in rest:
        for model, src, file_, n, ts in list_models(db):
            print(f"  {model:<16} [{src}] {n:>3} 行 | {file_}")
        return 0
    if "--model" in rest:
        # 按型号检索 (唯一索引): 三级树 / JSON / TSV, 支持 --sections 范围过滤与 --s9-clean 过滤无数据
        q = rest[rest.index("--model") + 1]
        hits = find_models(db, q)
        if not hits:
            print(f"  ✗ 库中无此型号: {q}")
            return 1
        mid, model, src, file_, n, ts = hits[0]
        sections = _parse_sections(rest)
        s9_active_only = ("--s9-clean" in rest or "--active-only" in rest or "--clean" in rest)
        out_path = rest[rest.index("--out") + 1] if ("--out" in rest and rest.index("--out") + 1 < len(rest)) else None
        if "--write-items" in rest:
            from core.msds_db import model_to_write_items
            items = model_to_write_items(db, mid, s9_active_only=True)
            import json as _json
            dumped = _json.dumps(items, ensure_ascii=False, indent=2)
            if out_path:
                Path(out_path).write_text(dumped, encoding="utf-8")
                print(f"✅ 覆写写入项已导出 → {out_path} (S9 有效字段 {len(items['sections'].get('9', []))} 项)")
            else:
                print(dumped)
            return 0
        if "--json" in rest:
            txt = render_model_json(db, mid, sections, s9_active_only=s9_active_only)
        elif "--tsv" in rest:
            txt = render_model_tsv(db, mid, sections, s9_active_only=s9_active_only)
        else:
            d = model_detail(db, mid)
            header = f"型号: {model} | 来源: {src} | 明细 {n} 行 | 入库: {ts}\n"
            if d.get("sha256"):
                header += f"sha256: {d['sha256'][:16]}… | 文件: {file_}\n"
            txt = header + render_model_tree(db, mid, sections, s9_active_only=s9_active_only)
        if out_path:
            enc = "utf-8-sig" if "--tsv" in rest else "utf-8"
            Path(out_path).write_text(txt, encoding=enc)
            print(f"✅ 已导出 → {out_path}")
        else:
            print(txt)
        return 0
    if "--model-search" in rest:
        # 关键词检索库内数据 → 命中型号清单 + 节/标签位置
        q = rest[rest.index("--model-search") + 1]
        sections = _parse_sections(rest)
        hits = model_search(db, q, sections)
        if not hits:
            print(f"  ✗ 无匹配: {q}")
            return 1
        by_model: dict[str, list] = {}
        for model, mid, sec, seq, label, value, kind, std_name in hits:
            by_model.setdefault(model, []).append((sec, seq, label, value, kind, std_name))
        for model, items in by_model.items():
            print(f"== {model} ({len(items)} 处命中) ==")
            for sec, seq, label, value, kind, std_name in items:
                tag = (std_name if (sec == 3 and kind == "component" and std_name) else (label or kind))
                print(f"   S{sec} {seq:<5} {tag:<30} {value}")
        return 0
    if "--query" in rest:
        q = rest[rest.index("--query") + 1]
        for model, sec, seq, label, value in search(db, q):
            print(f"  {model:<16} S{sec} {seq:<5} {label:<24} {value}")
        return 0
    if "--wide" in rest:
        model_name = rest[rest.index("--wide") + 1]
        row = db.execute(
            "SELECT model_id FROM msds_model WHERE model=? ORDER BY model_id LIMIT 1",
            (model_name,)).fetchone()
        if not row:
            print(f"  ✗ 未找到型号: {model_name}")
            return 1
        vals = wide_row(db, row[0])
        for k in sorted(vals):
            print(f"  {k:<34} = {vals[k][:80].replace(chr(10), ' / ')}")
        return 0
    if "--init" in rest:
        # 真正重建表结构 (schema 变更后使用): DROP msds_wide + 重灌 schema_field
        from core.msds_db import init_db
        init_db(db)
        print(f"  ✅ 表结构已重建: {db_path} (schema_field 字典 {db.execute('SELECT COUNT(*) FROM schema_field').fetchone()[0]} 项)")
        return 0

    # 入库
    xlsx_inputs: list[str] = []
    docx_inputs: list[str] = []
    i = 0
    while i < len(rest):
        if rest[i] == "--from-xlsx":
            xlsx_inputs.extend(rest[i + 1:])
            break
        docx_inputs.append(rest[i])
        i += 1

    if not docx_inputs and not xlsx_inputs:
        print(__doc__)
        return 2

    files = _canonical_docx_files(_collect_docx(docx_inputs))
    print(f"  docx {len(files)} 份 | xlsx 透视表 {len(xlsx_inputs)} 份 → {db_path}")

    from core.msds_db import wide_columns
    cols = wide_columns()

    failed = []
    for p in files:
        try:
            mid = insert_docx(db, p, cols)
            print(f"  ✓ {p.name} (model_id={mid})")
        except Exception as exc:
            failed.append((p.name, str(exc)))
            print(f"  ✗ {p.name}: {exc}")
    for x in xlsx_inputs:
        try:
            mid = insert_pivot_xlsx(db, x, cols)
            print(f"  ✓ xlsx {Path(x).name} (model_id={mid})")
        except Exception as exc:
            failed.append((Path(x).name, str(exc)))
            print(f"  ✗ xlsx {Path(x).name}: {exc}")

    # 兼容旧版 S9 查询表必须在整批入库结束后统一刷新，避免旧口径残留。
    refresh_s9_mapping(db)
    # 同步 Web 端型号父子级分类索引；分类表是运行时扩展，不改变冻结字段骨架。
    try:
        from build_model_categories import build as build_model_categories
        category_report = build_model_categories()
        print(f"  ✅ 型号父子级分类: 匹配 {category_report['matched']} / {category_report['models']}")
    except Exception as exc:
        print(f"  ⚠️ 型号父子级分类同步跳过: {exc}")

    n_model = db.execute("SELECT COUNT(*) FROM msds_model").fetchone()[0]
    n_field = db.execute("SELECT COUNT(*) FROM msds_field").fetchone()[0]
    n_wide = db.execute("SELECT COUNT(*) FROM msds_wide").fetchone()[0]
    print(f"\n  ✅ 入库完成: 型号 {n_model} | 明细 {n_field} 行 | 宽表 {n_wide} 行 | 失败 {len(failed)}")
    for name, err in failed:
        print(f"     - {name}: {err}")
    return 0 if not failed else 1


if __name__ == "__main__":
    sys.exit(main())
