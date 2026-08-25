# -*- coding: utf-8 -*-
"""数据库检索窗口 (库检索模式): 型号唯一索引 + 关键词检索, 复用现有呈现样式.

布局:
  - 顶栏: DB 路径 (可浏览) | 型号检索框 | 关键词检索框
  - 左侧: 型号列表 (型号/来源/行数, 点选进详查) + 17 节目录树
  - 右侧: SectionView (复用四列表格样式, show_rows 直接渲染库明细)
  - 状态栏: 型号 · 来源 · 入库时间 · 明细行数 · 命中提示

数据只读 SQLite (数据库/正式库/Data Base/msds_standard.db), 与 docx 直读检索独立.
"""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk, filedialog
import os
from pathlib import Path
import sys

from core.msds_db import (SEC_TITLES, find_models, listed_section_rows,
                          model_detail, model_search, model_to_write_items,
                          open_db)

from .section_tree import SectionView
from .theme import (COLOR_BORDER, COLOR_GRAY, COLOR_GREEN, COLOR_NAV,
                    COLOR_NAV_SEL, COLOR_PANEL, COLOR_ROW_ALT, COLOR_TEXT)

_PROJECT_ROOT = Path(os.environ.get(
    "MSDS_ROOT", str(Path(__file__).resolve().parents[2])))
DEFAULT_DB = _PROJECT_ROOT / "03-数据库" / "正式库" / "Data Base" / "msds_standard.db"
ENGINE_PROGRAM_DIR = _PROJECT_ROOT / "04-推断引擎" / "推断引擎程序"
if str(ENGINE_PROGRAM_DIR) not in sys.path:
    sys.path.insert(0, str(ENGINE_PROGRAM_DIR))
from pictogram_registry import display_value, extract_codes  # noqa: E402


class DbSearchWindow(tk.Toplevel):
    """库检索独立窗口 (工具栏「📚 数据库检索」打开)."""

    def __init__(self, master, db_path=None):
        super().__init__(master)
        self.title("📚 数据库检索 · MSDS 标准字段库")
        self.geometry("1240x760")
        self.configure(bg=COLOR_PANEL)
        self._conn = None
        self._current = None        # (model_id, model)
        self._current_model_name = ""
        self._current_sec = 1
        self._s9_clean_var = tk.BooleanVar(value=False)

        self._build_top()
        self._build_mid()
        self._build_status()

        if db_path:
            self._db_path_var.set(str(db_path))
        else:
            self._db_path_var.set(str(DEFAULT_DB))
        self._open_db()
        self.refresh_models("")

    # ---------------- 顶部 ----------------

    def _build_top(self):
        bar = tk.Frame(self, bg=COLOR_PANEL)
        bar.pack(fill="x", padx=8, pady=(8, 4))

        tk.Label(bar, text="📚 数据库检索", font=("Microsoft YaHei", 12, "bold"),
                 bg=COLOR_PANEL, fg=COLOR_TEXT).pack(side="left", padx=(0, 8))

        self._db_path_var = tk.StringVar()
        tk.Label(bar, text="库:", bg=COLOR_PANEL, fg=COLOR_GRAY).pack(side="left")
        e = tk.Entry(bar, textvariable=self._db_path_var, width=44)
        e.pack(side="left", padx=2)
        tk.Button(bar, text="浏览", command=self._browse_db,
                  bg=COLOR_NAV, fg="#fff", relief="flat").pack(side="left", padx=2)
        tk.Button(bar, text="重新打开", command=self._open_db,
                  bg=COLOR_NAV, fg="#fff", relief="flat").pack(side="left", padx=2)
        tk.Button(bar, text="📄 导出覆写写入项 (JSON)", command=self._export_write_items,
                  bg="#0B8043", fg="#fff", relief="flat", padx=10, pady=2,
                  font=("Microsoft YaHei", 9), cursor="hand2").pack(side="right", padx=(8, 0))

        row2 = tk.Frame(self, bg=COLOR_PANEL)
        row2.pack(fill="x", padx=8, pady=2)
        tk.Label(row2, text="🔍 型号:", bg=COLOR_PANEL, fg=COLOR_TEXT).pack(side="left")
        self._model_var = tk.StringVar()
        me = tk.Entry(row2, textvariable=self._model_var, width=22)
        me.pack(side="left", padx=2)
        me.bind("<Return>", lambda _e: self.refresh_models(self._model_var.get()))
        tk.Button(row2, text="检索型号", command=lambda: self.refresh_models(
            self._model_var.get()), bg=COLOR_NAV, fg="#fff", relief="flat").pack(side="left", padx=2)
        tk.Label(row2, text="关键词:", bg=COLOR_PANEL, fg=COLOR_TEXT).pack(side="left", padx=(12, 2))
        self._kw_var = tk.StringVar()
        ke = tk.Entry(row2, textvariable=self._kw_var, width=26)
        ke.pack(side="left", padx=2)
        ke.bind("<Return>", lambda _e: self._search_kw())
        tk.Button(row2, text="检索内容", command=self._search_kw,
                  bg=COLOR_GREEN, fg="#fff", relief="flat").pack(side="left", padx=2)
        tk.Button(row2, text="✕ 清除", command=self._clear,
                  bg=COLOR_BORDER, fg=COLOR_TEXT, relief="flat").pack(side="left", padx=2)

        # 过滤口子: S9 仅显有值项
        tk.Checkbutton(row2, text="S9仅显有值项 (非无数据)", variable=self._s9_clean_var,
                       command=self._on_toggle_s9_clean, bg=COLOR_PANEL, fg=COLOR_TEXT,
                       font=("Microsoft YaHei", 9), selectcolor=COLOR_PANEL).pack(side="left", padx=(14, 2))

    # ---------------- 中部: 型号列表 + 节树 + 视图 ----------------

    def _build_mid(self):
        mid = tk.Frame(self, bg=COLOR_PANEL)
        mid.pack(fill="both", expand=True, padx=8, pady=4)

        # 左栏: 型号列表
        left = tk.Frame(mid, bg=COLOR_PANEL)
        left.pack(side="left", fill="y")
        tk.Label(left, text="型号列表", font=("Microsoft YaHei", 10, "bold"),
                 bg=COLOR_PANEL, fg=COLOR_TEXT).pack(anchor="w")
        style = ttk.Style()
        style.configure("DbList.Treeview", rowheight=24, font=("Microsoft YaHei", 9))
        self.model_tree = ttk.Treeview(
            left, columns=("source", "n"), show="tree headings",
            style="DbList.Treeview", height=18)
        self.model_tree.heading("#0", text="型号")
        self.model_tree.heading("source", text="源")
        self.model_tree.heading("n", text="行")
        self.model_tree.column("#0", width=120)
        self.model_tree.column("source", width=40, anchor="center")
        self.model_tree.column("n", width=44, anchor="center")
        self.model_tree.pack(fill="both", expand=True, pady=(2, 6))
        self.model_tree.bind("<<TreeviewSelect>>", self._on_model_select)
        self._model_items: dict[str, int] = {}   # model -> model_id

        # 节目录
        tk.Label(left, text="节导航", font=("Microsoft YaHei", 10, "bold"),
                 bg=COLOR_PANEL, fg=COLOR_TEXT).pack(anchor="w")
        nstyle = ttk.Style()
        nstyle.configure("DbNav.Treeview", background=COLOR_NAV,
                         foreground="#E8EAED", fieldbackground=COLOR_NAV,
                         rowheight=26, borderwidth=0, font=("Microsoft YaHei", 9))
        nstyle.map("DbNav.Treeview", background=[("selected", COLOR_NAV_SEL)],
                   foreground=[("selected", "#FFFFFF")])
        self.sec_tree = ttk.Treeview(left, style="DbNav.Treeview", show="tree",
                                     selectmode="browse", height=12)
        self.sec_tree.pack(fill="both", expand=True)
        self.sec_tree.bind("<<TreeviewSelect>>", self._on_sec_select)

        # 右栏: 节内容视图
        self.view = SectionView(mid)
        self.view.pack(side="left", fill="both", expand=True, padx=(8, 0))

    def _build_status(self):
        self.status_var = tk.StringVar()
        st = tk.Label(self, textvariable=self.status_var, anchor="w",
                      bg=COLOR_ROW_ALT, fg=COLOR_TEXT,
                      font=("Microsoft YaHei", 9))
        st.pack(fill="x", side="bottom", padx=8, pady=6)

    # ---------------- 行为 ----------------

    def _open_db(self):
        p = self._db_path_var.get().strip()
        if not Path(p).exists():
            self.status_var.set(f"⚠️ 库文件不存在: {p}")
            return
        try:
            self._conn = open_db(p)
            self.status_var.set(f"已打开库: {p}")
        except Exception as exc:
            self.status_var.set(f"⚠️ 打开库失败: {exc}")

    def _browse_db(self):
        p = filedialog.askopenfilename(
            title="选择 MSDS 数据库", defaultextension=".db",
            filetypes=[("SQLite 数据库", "*.db"), ("所有文件", "*.*")])
        if p:
            self._db_path_var.set(p)
            self._open_db()

    def refresh_models(self, query: str):
        """型号列表 (支持按型号名过滤)."""
        if self._conn is None:
            return
        self.model_tree.delete(*self.model_tree.get_children())
        self._model_items.clear()
        hits = find_models(self._conn, query) if query.strip() else self._all_models()
        for mid, model, src, file_, n, ts in hits:
            iid = self.model_tree.insert("", "end", text=model,
                                         values=(src, n))
            self._model_items[iid] = mid
        if not query.strip():
            self.status_var.set(f"库内 {len(hits)} 条型号记录 — 点选型号进入详查")
        else:
            self.status_var.set(f"型号「{query}」命中 {len(hits)} 条")

    def _all_models(self):
        return self._conn.execute(
            "SELECT m.model_id, m.model, m.source, m.source_file,"
            " (SELECT COUNT(*) FROM msds_field f WHERE f.model_id=m.model_id),"
            " m.created_at FROM msds_model m ORDER BY m.model").fetchall()

    def _search_kw(self):
        """关键词检索 → 型号列表过滤为命中型号, 状态栏显示命中位置."""
        q = self._kw_var.get().strip()
        if not q or self._conn is None:
            return
        hits = model_search(self._conn, q)
        if not hits:
            self.status_var.set(f"✗ 无匹配: {q}")
            return
        # 命中型号集合 + 首条位置
        by_model: dict[str, tuple] = {}
        for model, mid, sec, seq, label, value, kind in hits:
            by_model.setdefault(model, (mid, []))
            by_model[model][1].append((sec, seq, label, value, kind))
        self.model_tree.delete(*self.model_tree.get_children())
        self._model_items.clear()
        for model, (mid, items) in by_model.items():
            n = len(items)
            secs = ",".join(str(s[0]) for s in items[:5])
            iid = self.model_tree.insert(
                "", "end", text=model, values=("命中", n))
            self._model_items[iid] = mid
        first = list(by_model.items())[0]
        self.status_var.set(
            f"关键词「{q}」命中 {len(by_model)} 个型号 / 共 {len(hits)} 处 "
            f"— 如 {first[0]} S{first[1][1][0][0]} {first[1][1][0][1]}: "
            f"{first[1][1][0][3][:30]}")

    def _clear(self):
        self._model_var.set("")
        self._kw_var.set("")
        self.refresh_models("")

    def _on_toggle_s9_clean(self):
        """切换 S9 是否仅显示有值项."""
        if self._current is not None:
            self._render_section(self._current_sec)

    def _export_write_items(self):
        """导出当前型号的纯净覆写写入项 JSON (S9 仅有值项, 直通覆写模块)."""
        if self._conn is None or self._current is None:
            self.status_var.set("⚠️ 请先在左侧选择一个型号")
            return
        mid, model = self._current
        items = model_to_write_items(self._conn, mid, s9_active_only=True)
        import json
        default_name = f"write_items_{model}_pure.json"
        p = filedialog.asksaveasfilename(
            title=f"导出 {model} 覆写写入项 (JSON)",
            initialfile=default_name,
            defaultextension=".json",
            filetypes=[("JSON 文件", "*.json"), ("所有文件", "*.*")])
        if p:
            Path(p).write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")
            s9_n = len(items.get("sections", {}).get("9", []))
            self.status_var.set(f"✅ 已导出 {model} 写入项 → {Path(p).name} (S9 有效字段 {s9_n} 项)")

    def _on_model_select(self, _event):
        sel = self.model_tree.selection()
        if not sel:
            return
        mid = self._model_items.get(sel[0])
        if mid is None:
            return
        self._load_model(mid)

    def _load_model(self, model_id: int):
        d = model_detail(self._conn, model_id)
        if not d:
            return
        self._current = (model_id, d["model"])
        self._current_model_name = d["model"]
        # 节导航: 固定 17 节 (只显示库中存在的节)
        self.sec_tree.delete(*self.sec_tree.get_children())
        self._sec_items: dict[int, str] = {}
        have = {r[0] for r in self._conn.execute(
            "SELECT DISTINCT section FROM msds_field WHERE model_id=?",
            (model_id,))}
        for num, title in sorted(SEC_TITLES.items()):
            if num not in have:
                continue
            iid = self.sec_tree.insert("", "end", text=f"{num}. {title}")
            self._sec_items[num] = iid
        self.status_var.set(
            f"{d['model']} [{d['source']}] 明细 {d['fields_count']} 行 · "
            f"入库 {d['created_at']} · 文件 {d['source_file']}")
        # 默认渲染第一个有数据的节
        first = sorted(have)[0] if have else 1
        self._current_sec = first
        self.sec_tree.selection_set(self._sec_items.get(first, ""))
        self._render_section(first)

    def _on_sec_select(self, _event):
        sel = self.sec_tree.selection()
        if not sel:
            return
        for num, iid in self._sec_items.items():
            if iid == sel[0]:
                self._current_sec = num
                self._render_section(num)
                break

    def _render_section(self, num: int):
        if self._current is None:
            return
        self._current_sec = num
        mid, model = self._current
        d = model_detail(self._conn, mid)
        title = f"{num}. {SEC_TITLES.get(num, f'第{num}节')}"
        s9_clean = self._s9_clean_var.get()
        # 清单骨架渲染: 结构与 PEA-4139 模板参照一致, 缺值标「无数据」; s9_clean=True 时仅显有效数据项
        rows = listed_section_rows(self._conn, mid, num, s9_active_only=s9_clean)
        if num == 2:
            self._augment_s2_pictogram_candidates(rows, model)
        mode_tag = " (S9仅有值项)" if (num == 9 and s9_clean) else ""
        meta = f"{model} [{d.get('source', '')}] · 第{num}节 · 按飞书清单结构渲染{mode_tag}"
        self.view.show_rows(num, f"{title} — {model}", rows, meta=meta)

    @staticmethod
    def _augment_s2_pictogram_candidates(rows, model: str) -> None:
        """只读显示层补足 S2 象形图候选，不把候选结论写回正式库。

        历史型号的 S2 常只保存 ``[象形图]`` 占位，检索 GUI 无法从占位值恢复
        GHS 代码。这里复用 Section 2 事实包生成候选代码；候选值明确带状态，
        只有通过人工审核/后续写入流程后，才可作为正式 MSDS 结论。
        """
        labels = {"象形图", "GHS象形图", "GHS-象形图"}
        if not any(getattr(row, "kind", "") == "field"
                   and str(getattr(row, "label", "")) in labels
                   for row in rows):
            return
        try:
            from s2_fact_package import collect_fact_package
            package = collect_fact_package(model)
            inference = package.get("inference") or {}
            candidate = inference.get("candidate_label_elements") or {}
            codes = extract_codes(candidate.get("pictograms") or "")
            if not codes:
                codes = list(candidate.get("pictogram_codes") or [])
            if not codes:
                return
            status = inference.get("status") or "manual_review"
            display = display_value(codes)
        except Exception:
            # 检索 GUI 不能因为推断源缺失而阻断原有只读查询。
            return
        for row in rows:
            if getattr(row, "kind", "") != "field" or str(getattr(row, "label", "")) not in labels:
                continue
            current = str(getattr(row, "value", "") or "").strip()
            if extract_codes(current):
                continue
            if current and not current.startswith("[象形图") and current not in {"无", "无数据"}:
                continue
            row.value = f"{display}（推断候选，状态：{status}）"
