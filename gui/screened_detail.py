# -*- coding: utf-8 -*-
"""正式筛选库右侧详情: 信息栏 + 16节切换条 + 连续表格 (保留原 Word 表格形式).

与原「16 节结构」一致:
  - 整节渲染为一张连续表格: 外框深蓝 2px + 行分隔线 + 交替行底色
  - 节标题 = 通栏深蓝表头行
  - 字段行 = 标题列(锁定) | 值列(自动换行)
统一库型号: 按 16 节分组渲染; 宽表型号: 同表格, 列=标准标题|原始标题|值|品牌.
"""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from .theme import (
    COLOR_BORDER, COLOR_GRAY, COLOR_NAV, COLOR_PANEL, COLOR_ROW_ALT, COLOR_TEXT,
)


class TableSheet(ttk.Frame):
    """可滚动连续表格画布 (保留原表格形态)."""

    def __init__(self, master):
        super().__init__(master)
        self.canvas = tk.Canvas(self, bg=COLOR_PANEL, highlightthickness=0)
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.inner = tk.Frame(self.canvas, bg=COLOR_PANEL)
        self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(
            scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", self._on_canvas_cfg)
        self.canvas.bind("<MouseWheel>",
                         lambda e: self.canvas.yview_scroll(int(-e.delta / 120), "units"))

        self._table = None
        self._ri = 0
        self._labels = []          # 参与 wrap 自适应的 Label

    # ---------- 画布/滚动 ----------

    def _on_canvas_cfg(self, e):
        self.canvas.itemconfigure(self._win, width=e.width)
        # 值列(最后一列)弹性换行, 标题列固定
        wrap = max(160, e.width - 320)
        for l in self._labels:
            l.configure(wraplength=wrap)

    def clear(self):
        for w in self.inner.winfo_children():
            w.destroy()
        self._labels = []

    # ---------- 表格容器 ----------

    def begin(self, ncols=2):
        """创建整张连续表格: 外框 2px 深蓝 + 内框线底色."""
        outer = tk.Frame(self.inner, bg=COLOR_BORDER, highlightthickness=2,
                         highlightbackground=COLOR_NAV)
        outer.pack(fill="x", pady=(0, 8))
        self._table = tk.Frame(outer, bg=COLOR_BORDER)
        self._table.pack(fill="both", expand=True)
        self._ri = 0
        self._ncols = ncols
        # 列权重: 第0列标题固定, 其余列弹性
        for ci in range(ncols):
            self._table.grid_columnconfigure(ci, weight=(1 if ci >= 1 else 0))

    def title(self, text):
        """通栏节标题行 (深蓝表头, 跨全宽)."""
        cell = tk.Frame(self._table, bg=COLOR_NAV)
        cell.grid(row=self._ri, column=0, columnspan=self._ncols, sticky="nsew",
                  padx=(0, 0), pady=(0, 1))
        tk.Label(cell, text=text, bg=COLOR_NAV, fg="#FFFFFF",
                 font=("Microsoft YaHei", 11, "bold"), anchor="w",
                 padx=12, pady=6).pack(fill="x")
        self._ri += 1

    def header(self, cells):
        """表头行 (深蓝底, 用于宽表多列表头)."""
        for ci, text in enumerate(cells):
            cell = tk.Frame(self._table, bg=COLOR_NAV)
            cell.grid(row=self._ri, column=ci, sticky="nsew",
                      padx=(0, 0 if ci == len(cells) - 1 else 1), pady=(0, 1))
            cell.grid_rowconfigure(0, weight=1)
            tk.Label(cell, text=text, bg=COLOR_NAV, fg="#FFFFFF",
                     font=("Microsoft YaHei", 10, "bold"), anchor="w",
                     padx=10, pady=5).grid(row=0, column=0, sticky="nsew")
        self._ri += 1

    def row(self, cells, wraps=()):
        """字段行: cells 逐列渲染, 交替底色 + 行间/列间 1px 分隔线.
        wraps: 需要随窗口换行的列索引 (默认最后一列)."""
        bg = COLOR_ROW_ALT if (self._ri % 2 == 1) else COLOR_PANEL
        n = len(cells)
        for ci, text in enumerate(cells):
            cell = tk.Frame(self._table, bg=bg)
            cell.grid(row=self._ri, column=ci, sticky="nsew",
                      padx=(0, 0 if ci == n - 1 else 1), pady=(0, 1))
            cell.grid_rowconfigure(0, weight=1)
            cell.grid_columnconfigure(0, weight=1)
            l = tk.Label(cell, text=str(text), bg=bg, fg=COLOR_TEXT,
                         font=("Microsoft YaHei", 10, "bold" if ci == 0 else "normal"),
                         anchor="nw", justify="left", wraplength=0, padx=10, pady=6)
            l.grid(row=0, column=0, sticky="nsew")
            if ci in wraps or ci == n - 1:
                self._labels.append(l)
        self._ri += 1


class ScreenedDetail(ttk.Frame):
    """右侧详情: 信息栏 + 16节切换条 + 连续表格."""

    def __init__(self, master, on_status=None):
        super().__init__(master)
        self.on_status = on_status
        self._kind = "unified"   # unified | wide
        self._model = ""
        self._sections: dict[int, list] = {}

        # 信息栏
        info = tk.Frame(self, bg=COLOR_PANEL, height=40)
        info.pack(fill="x", padx=2, pady=(0, 2))
        info.pack_propagate(False)
        self.info_var = tk.StringVar(value="请选择左侧节点")
        tk.Label(info, textvariable=self.info_var, bg=COLOR_PANEL, fg=COLOR_TEXT,
                 font=("Microsoft YaHei", 11, "bold"), anchor="w").pack(side="left", padx=10)
        self.meta_var = tk.StringVar(value="")
        tk.Label(info, textvariable=self.meta_var, bg=COLOR_PANEL, fg=COLOR_GRAY,
                 font=("Microsoft YaHei", 9), anchor="w").pack(side="left", padx=14)

        # 16 节切换条
        bar = tk.Frame(self, bg="#EDF0F5")
        bar.pack(fill="x", padx=2, pady=(0, 2))
        self._sec_btns: list[tk.Button] = []
        for n in range(1, 17):
            b = tk.Button(bar, text=str(n), width=3, relief="flat", bg="#FFFFFF",
                          fg=COLOR_TEXT, font=("Microsoft YaHei", 9), cursor="hand2",
                          command=lambda num=n: self._click_section(num))
            b.pack(side="left", padx=1, pady=3)
            self._sec_btns.append(b)
        tk.Label(bar, text="节", bg="#EDF0F5", fg=COLOR_GRAY,
                 font=("Microsoft YaHei", 9)).pack(side="left", padx=(4, 0))

        # 连续表格
        self.sheet = TableSheet(self)
        self.sheet.pack(fill="both", expand=True, padx=2, pady=(0, 2))

    # ---------- 数据进入 ----------

    def show_model(self, model: str, meta: str, sections: dict[int, list]):
        """统一库型号. sections: {节号: [(标题, 值), ...]}"""
        self._kind = "unified"
        self._model = model
        self._sections = sections
        self.info_var.set(f"📦 {model}")
        self.meta_var.set(meta)
        first = min(sections.keys()) if sections else 1
        self._render_section(first)

    def show_wide(self, model: str, meta: str, sections: dict[int, list]):
        """宽表型号. sections: {节号: [(标准标题, 原始标题, 值, 品牌), ...]}"""
        self._kind = "wide"
        self._model = model
        self._sections = sections
        self.info_var.set(f"📋 {model}")
        self.meta_var.set(meta)
        first = min(sections.keys()) if sections else 1
        self._render_wide_section(first)

    # ---------- 渲染 ----------

    def _click_section(self, num: int):
        if self._kind == "wide":
            self._render_wide_section(num)
        else:
            self._render_section(num)

    def _highlight(self, num: int):
        for i, b in enumerate(self._sec_btns, start=1):
            if i == num:
                b.configure(bg=COLOR_NAV, fg="#FFFFFF")
            else:
                b.configure(bg="#FFFFFF", fg=COLOR_TEXT)

    def _render_section(self, num: int):
        self._highlight(num)
        self.sheet.clear()
        self.sheet.begin(ncols=2)
        self.sheet.title(f"第{num}节")
        rows = self._sections.get(num, [])
        if not rows:
            self.sheet.title("(本节无字段)")
            return
        for label, value in rows:
            self.sheet.row([label, value], wraps=(1,))
        # 刷新值列换行 (画布宽已知后)
        w = self.sheet.canvas.winfo_width()
        if w > 1:
            self.sheet._on_canvas_cfg(type("E", (), {"width": w}))

    def _render_wide_section(self, num: int):
        self._highlight(num)
        self.sheet.clear()
        self.sheet.begin(ncols=4)
        self.sheet.title(f"第{num}节")
        rows = self._sections.get(num, [])
        if not rows:
            self.sheet.title("(本节无记录)")
            return
        self.sheet.header(["标准标题", "原始标题", "值", "品牌"])
        for std, raw, value, vendor in rows:
            self.sheet.row([std, raw, value, vendor], wraps=(2,))
        w = self.sheet.canvas.winfo_width()
        if w > 1:
            self.sheet._on_canvas_cfg(type("E", (), {"width": w}))