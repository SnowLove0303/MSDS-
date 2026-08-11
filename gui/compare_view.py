# -*- coding: utf-8 -*-
"""覆写指向比对视图: 模板字段 vs 产品字段 差异表."""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from core.compare import CompareResult

from .theme import (
    COLOR_GRAY, COLOR_GREEN, COLOR_ACCENT, COLOR_ORANGE, COLOR_RED,
    COLOR_PANEL, COLOR_TEXT, COLOR_YELLOW, WRITE_COLORS, WRITE_LABELS,
)


class CompareView(ttk.Frame):
    """比对结果: 字段级覆写指向表."""

    def __init__(self, master):
        super().__init__(master, padding=6)
        self._build()

    def _build(self):
        self.title_var = tk.StringVar()
        tk.Label(self, textvariable=self.title_var, font=("Microsoft YaHei", 13, "bold"),
                 bg=COLOR_PANEL, fg=COLOR_TEXT, anchor="w").pack(fill="x", pady=(0, 6))

        # 概览统计条
        self.summary_var = tk.StringVar()
        tk.Label(self, textvariable=self.summary_var, font=("Microsoft YaHei", 10),
                 bg=COLOR_PANEL, fg=COLOR_TEXT, anchor="w").pack(fill="x", pady=(0, 6))

        cols = ("section", "label", "template", "product", "action", "reason")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=16)
        self.tree.heading("section", text="节")
        self.tree.heading("label", text="字段")
        self.tree.heading("template", text="模板现值")
        self.tree.heading("product", text="产品提供值")
        self.tree.heading("action", text="覆写指向")
        self.tree.heading("reason", text="原因")
        self.tree.column("section", width=40, anchor="center", stretch=False)
        self.tree.column("label", width=200, anchor="w")
        self.tree.column("template", width=180, anchor="w")
        self.tree.column("product", width=180, anchor="w")
        self.tree.column("action", width=90, anchor="center", stretch=False)
        self.tree.column("reason", width=300, anchor="w")

        # 配色
        style = ttk.Style()
        style.configure("Cmp.Treeview", rowheight=26, font=("Microsoft YaHei", 9))
        for ws, color in WRITE_COLORS.items():
            style.map(f"Cmp.{ws}.Treeview",
                      background=[("selected", "#E8F0FE")],
                      foreground=[("selected", color)])
            style.configure(f"Cmp.{ws}.Treeview", foreground=color)

        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.residue_label = tk.Label(self, text="", font=("Microsoft YaHei", 10),
                                      bg=COLOR_YELLOW, fg="#5A3E00", anchor="w", justify="left",
                                      wraplength=900)
        self.residue_label.pack(side="bottom", fill="x", pady=(6, 0))

    def show_compare(self, cr: CompareResult):
        self.title_var.set(f"覆写指向分析: {cr.template_file} ← {cr.product_file}")
        counts = cr.counts()
        total = len(cr.decisions)
        self.summary_var.set(
            f"共 {total} 字段 | 保留模板 {counts.get('template', 0)} | "
            f"产品覆盖 {counts.get('product', 0)} | 清空 {counts.get('clear', 0)} | "
            f"新增 {counts.get('add', 0)} | 需确认 {counts.get('review', 0)}"
        )

        self.tree.delete(*self.tree.get_children())
        for d in cr.decisions:
            color = WRITE_COLORS.get(d.write_source, COLOR_GRAY)
            self.tree.insert("", "end", values=(
                d.section, d.label,
                d.template_value[:40], d.product_value[:40],
                WRITE_LABELS.get(d.write_source, d.write_source),
                d.reason,
            ), tags=(d.write_source,))
        for ws in WRITE_COLORS:
            self.tree.tag_configure(ws, foreground=WRITE_COLORS[ws])

        if cr.residue_risks:
            lines = [f"⚠️ {len(cr.residue_risks)} 处模板残留旧值风险: " +
                     "、".join(f"S{d.section}「{d.label}」" for d in cr.residue_risks[:8])]
            if len(cr.residue_risks) > 8:
                lines[0] += f" 等{len(cr.residue_risks)}处"
            self.residue_label.configure(text=lines[0])
        else:
            self.residue_label.configure(text="✅ 无模板残留旧值风险")
