# -*- coding: utf-8 -*-
"""MSDS 结构读取程序入口.

用法:
  python main.py                        # 启动 GUI
  python main.py --cli <文件>           # 命令行解析并打印 16 节摘要
  python main.py --extract <文件>       # 分层检索提取 (section→大标题→小标题→字段)
         [--query 关键词] [--scope label|value|all|section]
         [--json|--tsv] [--out 文件]
  python main.py --extract-many <目录或文件...> --sections 1,3,9   # 批量提取指定节
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

from core.docx_reader import read_msds
from core.extract import (extract_doc, extract_many, export_tsv, print_hierarchy,
                          render_json, render_tsv)


def cli(path: str) -> int:
    result = read_msds(path)
    print(f"文件: {result.file_name}")
    print(f"摘要: {result.summary()}")
    if result.header:
        print(f"页眉: {result.header!r}")
    if result.footer:
        print(f"页脚: {result.footer!r}")
    for a in result.anomalies:
        print(f"  [{'⚠️' if a.level == 'warn' else '❌'}] S{a.section}: {a.message} {a.detail}")
    print()
    for n in sorted(result.sections):
        sec = result.sections[n]
        print(f"=== {sec.full_title} ===")
        for f in sec.fields:
            print(f"  {f.label}: {f.value}")
        for c in sec.components:
            print(f"  成分: {c.name} | {c.cas} | {c.conc}")
        for ln in sec.lines:
            print(f"  · {ln}")
    return 0


def _extract_cli(argv: list[str]) -> int:
    """--extract 分层检索提取 (可批量, 多文件)."""
    paths: list[str] = []
    query: str | None = None
    scope = "all"
    fmt = "text"          # text(三级树) | json(嵌套树) | tsv(扁平)
    out: str | None = None
    sections: set[int] | None = None
    flat = False          # --flat: 扁平输出 (旧版行为)
    i = 0
    while i < len(argv):
        a = argv[i]
        if a in ("--query", "-q") and i + 1 < len(argv):
            query = argv[i + 1]; i += 2
        elif a == "--scope" and i + 1 < len(argv):
            scope = argv[i + 1]; i += 2
        elif a in ("--json", "-j"):
            fmt = "json"; i += 1
        elif a in ("--tsv", "-t"):
            fmt = "tsv"; i += 1
        elif a == "--flat":
            flat = True; i += 1
        elif a == "--out" and i + 1 < len(argv):
            out = argv[i + 1]; i += 2
        elif a == "--sections" and i + 1 < len(argv):
            sections = {int(x) for x in argv[i + 1].split(",") if x.strip().isdigit()}
            i += 2
        else:
            paths.append(a); i += 1
    if not paths:
        print("用法: python main.py --extract <docx...> [--query 词] [--scope label|value|all|section] [--json|--tsv] [--flat] [--out 文件] [--sections 1,3,9]")
        return 2
    if len(paths) == 1:
        r = read_msds(paths[0])
        # 三级父子级树 (对应 GUI 表格: 节 → 大标题 → 字段)
        from core.extract import (build_hierarchy, flatten_nodes, render_tree,
                                  render_tree_json, search_tree)
        nodes = build_hierarchy(r)
        if sections:
            nodes = [n for n in nodes if n.number in sections]
        if query:
            nodes = search_tree(nodes, query, scope)
        entries = flatten_nodes(nodes)   # 统计用
        if fmt == "json":
            text = render_tree_json(nodes)
        elif fmt == "tsv":
            from core.extract import render_tsv
            text = render_tsv(entries)
        elif flat:
            text = f"文件: {r.file_name} | 共 {len(entries)} 条\n" + _text_of(entries)
        else:
            text = f"文件: {r.file_name} | 共 {len(entries)} 条 (三级父子级)\n" + render_tree(nodes)
        if out:
            enc = "utf-8-sig" if fmt == "tsv" else "utf-8"  # TSV 需 BOM(Excel), JSON 纯utf-8
            Path(out).write_text(text, encoding=enc)
            print(f"✅ 已导出 {len(entries)} 条 → {out}")
        else:
            print(text)
        return 0
    # 批量: 多文件 → 每文件三级树
    from core.extract import (build_hierarchy, flatten_nodes, render_tree,
                              render_tree_json, search_tree)
    collected: dict[str, list] = {}
    total = 0
    for p in paths:
        try:
            r = read_msds(p)
            nodes = build_hierarchy(r)
            if sections:
                nodes = [n for n in nodes if n.number in sections]
            if query:
                nodes = search_tree(nodes, query, scope)
            collected[Path(p).name] = nodes
            total += len(flatten_nodes(nodes))
        except Exception as exc:
            collected[Path(p).name] = []
            print(f"⚠️ 读取失败 {p}: {exc}")
    if fmt == "json":
        text = json.dumps({k: [n.to_dict() for n in v] for k, v in collected.items()},
                          ensure_ascii=False, indent=2)
    elif fmt == "tsv":
        rows = ["文件\t节\t大标题\t小标题\t标签\t内容"]
        for fname, nodes in collected.items():
            for e in flatten_nodes(nodes):
                rows.append("\t".join([fname, str(e.section), e.big_title, e.sub_title,
                                       e.full_label(), e.value.replace("\t", " ").replace("\n", " ")]))
        text = "\n".join(rows)
    elif flat:
        buf = [f"共 {len(paths)} 文件 / {total} 条"]
        for fname, nodes in collected.items():
            buf.append(f"\n===== {fname} ({len(flatten_nodes(nodes))} 条) =====")
            buf.append(_text_of(flatten_nodes(nodes)))
        text = "\n".join(buf)
    else:
        buf = [f"共 {len(paths)} 文件 / {total} 条 (三级父子级)"]
        for fname, nodes in collected.items():
            buf.append(f"\n===== {fname} ({len(flatten_nodes(nodes))} 条) =====")
            buf.append(render_tree(nodes))
        text = "\n".join(buf)
    if out:
        enc = "utf-8-sig" if fmt == "tsv" else "utf-8"
        Path(out).write_text(text, encoding=enc)
        print(f"✅ 已导出 {total} 条 / {len(paths)} 文件 → {out}")
    else:
        print(text)
    return 0


def _text_of(entries) -> str:
    """分层文本 (复用 core.extract.render_text 的行, 但去掉首行统计)."""
    from core.extract import render_text
    return render_text(entries)


def main() -> int:
    if len(sys.argv) > 1 and sys.argv[1] in ("--extract", "-e"):
        return _extract_cli(sys.argv[2:])
    if len(sys.argv) > 1 and sys.argv[1] == "--cli":
        if len(sys.argv) < 3:
            print("用法: python main.py --cli <docx路径>")
            return 2
        return cli(sys.argv[2])
    if len(sys.argv) > 1:
        return cli(sys.argv[1])

    try:
        from gui.main_window import MainWindow
    except ImportError as exc:
        print(f"GUI 依赖缺失: {exc}")
        print("请确认 python-docx 与 tkinter 已安装 (python -c 'import docx, tkinter')")
        return 1

    app = MainWindow()
    app.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
