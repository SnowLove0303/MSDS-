# -*- coding: utf-8 -*-
"""表单系统窗口: 录入 Section 1 / 3 / 9 输入结构 (参照 PEA-4139 表单).

使用者在表单里填写产品/供应商 (S1)、产品类型与成分 (S3)、基础物性 (S9),
点击「生成写入项」→ 导出 {S1,S3,S9} 写入项 JSON (供覆写引擎 / 后续推断引擎).

与主窗口是独立子窗口: 主窗口工具栏「表单系统」打开本窗口.
"""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import tkinter as tk

from core.form_schema import (
    ComponentRow, S1_FIELDS, S3_FIELDS, S9_FIELDS, build_form_schema,
    form_to_write_items,
)
from gui.theme import (
    COLOR_BG, COLOR_BORDER, COLOR_GREEN, COLOR_NAV, COLOR_PANEL, COLOR_TEXT,
)


class S139FormPanel(tk.Frame):
    """可嵌入的 S1/S3/S9 表单面板。

    这是表单系统的唯一输入实现。独立表单窗口和推断工作台都复用本面板，
    保证字段骨架、成分行、S9 动态行和切换保存逻辑一致。
    """

    def __init__(self, master, *, show_actions=True, on_close=None):
        super().__init__(master, bg=COLOR_BG)
        self._show_actions = show_actions
        self._on_close = on_close

        self._schema = build_form_schema()

        # ---- 数据模型 (权威状态, 不随 widget 销毁) ----
        self.s1_data = {f.label: "" for f in S1_FIELDS}
        self.s3_product_type = ""
        self.s3_components = [["", "", ""] for _ in range(3)]   # [name, cas, conc]
        self.s9_data = [{"label": f.label, "value": ""} for f in S9_FIELDS]

        # ---- 当前节 widget 层 ----
        self._cur_section = None           # 当前节标识: 1/3/9
        self._s1_vars = {}                 # label -> StringVar
        self._comp_row_vars = []           # [ [var_name, var_cas, var_conc], ... ]
        self._s9_row_vars = []             # [ {"seq_label", "vlabel", "vvalue"}, ... ]
        self.product_type_var = tk.StringVar()
        self._comp_container = None
        self._s9_container = None
        self._form_canvas = None
        self._form_inner = None

        self._build_nav()
        self._build_body()
        self._show_section("1")

    def _build_nav(self):
        nav = tk.Frame(self, bg=COLOR_NAV, width=190)
        nav.pack(side="left", fill="y")
        nav.pack_propagate(False)
        tk.Label(nav, text="表单节", bg=COLOR_NAV, fg="#E8EAED",
                 font=("Microsoft YaHei", 12, "bold"), padx=16, pady=12).pack(fill="x")
        for key, label in (("1", "S1 物料及供应商标识"),
                           ("3", "S3 成分/组成资料"),
                           ("9", "S9 物理和化学特性")):
            b = tk.Button(nav, text=label, bg=COLOR_NAV, fg="#E8EAED",
                          relief="flat", anchor="w", padx=16, pady=8,
                          font=("Microsoft YaHei", 10), cursor="hand2",
                          command=lambda k=key: self._show_section(k))
            b.pack(fill="x", padx=6, pady=2)
        if self._show_actions:
            act = tk.Frame(nav, bg=COLOR_NAV)
            act.pack(side="bottom", fill="x", padx=8, pady=10)
            tk.Button(act, text="生成写入项 JSON", command=self._export,
                      bg=COLOR_GREEN, fg="white", relief="flat", padx=8, pady=6,
                      font=("Microsoft YaHei", 10, "bold"), cursor="hand2").pack(fill="x")
            tk.Button(act, text="关闭",
                      command=self._on_close or self.destroy,
                      bg="#E8EAED", fg=COLOR_TEXT, relief="flat", padx=8, pady=4,
                      font=("Microsoft YaHei", 9)).pack(fill="x", pady=(6, 0))

    def _build_body(self):
        body = tk.Frame(self, bg=COLOR_BG)
        body.pack(side="left", fill="both", expand=True, padx=(0, 6), pady=6)
        self._form_canvas = tk.Canvas(body, bg=COLOR_PANEL, highlightthickness=1,
                                      highlightbackground=COLOR_BORDER)
        vsb = ttk.Scrollbar(body, orient="vertical", command=self._form_canvas.yview)
        self._form_canvas.configure(yscrollcommand=vsb.set)
        self._form_canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self._form_inner = tk.Frame(self._form_canvas, bg=COLOR_PANEL)
        self._win = self._form_canvas.create_window((0, 0), window=self._form_inner, anchor="nw")
        self._form_inner.bind("<Configure>",
                              lambda _e: self._form_canvas.configure(scrollregion=self._form_canvas.bbox("all")))
        self._form_canvas.bind("<Configure>",
                               lambda e: self._form_canvas.itemconfigure(self._win, width=e.width))
        self._form_canvas.bind("<MouseWheel>",
                               lambda e: self._form_canvas.yview_scroll(int(-e.delta / 120), "units"))

    def _clear_form(self):
        for w in self._form_inner.winfo_children():
            w.destroy()

    def _show_section(self, sec):
        if sec == self._cur_section:
            return                     # 已在目标节, 不重复重建
        self._commit_current()         # 先把当前节 widget 值写回数据模型
        self._clear_form()
        self._cur_section = sec
        if sec == "1":
            self._render_s1()
        elif sec == "3":
            self._render_s3()
        elif sec == "9":
            self._render_s9()

    def _render_current(self):
        """按当前节重建视图 (增删行后恢复)."""
        sec = self._cur_section
        self._clear_form()
        if sec == "1":
            self._render_s1()
        elif sec == "3":
            self._render_s3()
        elif sec == "9":
            self._render_s9()

    def _commit_current(self):
        """把当前节 widget 值写回数据模型 (切换/导出前调用)."""
        if self._cur_section == "1":
            for label, var in self._s1_vars.items():
                self.s1_data[label] = var.get()
        elif self._cur_section == "3":
            self.s3_product_type = self.product_type_var.get()
            self.s3_components = [[v.get() for v in row] for row in self._comp_row_vars]
        elif self._cur_section == "9":
            self.s9_data = [{"label": r["vlabel"].get(), "value": r["vvalue"].get()}
                            for r in self._s9_row_vars]

    def get_data(self):
        """提交当前控件并返回 S1/S3/S9 原始表单数据。"""
        self._commit_current()
        return {
            "s1": dict(self.s1_data),
            "s3_product_type": self.s3_product_type,
            "s3_components": [list(row) for row in self.s3_components],
            "s9": [dict(row) for row in self.s9_data],
        }

    def reset_data(self):
        """清空表单并回到 S1。"""
        self.s1_data = {f.label: "" for f in S1_FIELDS}
        self.s3_product_type = ""
        self.s3_components = [["", "", ""] for _ in range(3)]
        self.s9_data = [{"label": f.label, "value": ""} for f in S9_FIELDS]
        self._cur_section = None
        self._show_section("1")

    def _section_header(self, title, subtitle=""):
        tk.Label(self._form_inner, text=title, bg=COLOR_NAV, fg="#FFFFFF",
                 font=("Microsoft YaHei", 12, "bold"), anchor="w",
                 padx=14, pady=8).pack(fill="x")
        if subtitle:
            tk.Label(self._form_inner, text=subtitle, bg=COLOR_PANEL, fg="#5F6368",
                     font=("Microsoft YaHei", 9), anchor="w", padx=14,
                     pady=2).pack(fill="x", pady=4)

    def _render_s1(self):
        self._section_header("S1 物料及供应商标识",
                             "使用者填写产品名称 / 供应商信息（参照 PEA-4139 表单）")
        self._s1_vars.clear()
        for f in S1_FIELDS:
            row = tk.Frame(self._form_inner, bg=COLOR_PANEL)
            row.pack(fill="x", padx=14, pady=3)
            tk.Label(row, text=f.label, bg=COLOR_PANEL, fg=COLOR_TEXT,
                     font=("Microsoft YaHei", 10), width=24, anchor="w").pack(side="left")
            var = tk.StringVar(value=self.s1_data.get(f.label, ""))
            e = tk.Entry(row, textvariable=var, font=("Microsoft YaHei", 10), relief="solid",
                         highlightthickness=1, highlightbackground=COLOR_BORDER)
            e.pack(side="left", fill="x", expand=True, ipady=2)
            if f.seq:
                tk.Label(row, text=f.seq, bg=COLOR_PANEL, fg="#80868B",
                         font=("Microsoft YaHei", 9), width=7, anchor="e").pack(side="left", padx=(6, 2))
            self._s1_vars[f.label] = var

    def _render_s3(self):
        self._section_header("S3 成分/组成资料",
                             "产品类型 + 成分表（化学品名称 + CAS编号 + 含量%）")
        row = tk.Frame(self._form_inner, bg=COLOR_PANEL)
        row.pack(fill="x", padx=14, pady=3)
        tk.Label(row, text="产品类型 *", bg=COLOR_PANEL, fg=COLOR_TEXT,
                 font=("Microsoft YaHei", 10), width=24, anchor="w").pack(side="left")
        self.product_type_var.set(self.s3_product_type)
        cb = ttk.Combobox(row, textvariable=self.product_type_var, state="readonly",
                          values=S3_FIELDS[0].options, font=("Microsoft YaHei", 10))
        cb.pack(side="left", fill="x", expand=True, ipady=1)
        tk.Label(row, text="3.1", bg=COLOR_PANEL, fg="#80868B",
                 font=("Microsoft YaHei", 9), width=7, anchor="e").pack(side="left", padx=(6, 2))
        hdr = tk.Frame(self._form_inner, bg=COLOR_NAV)
        hdr.pack(fill="x", padx=14, pady=(12, 2))
        for t, w in (("化学品名称", 3), ("CAS编号", 2), ("含量%（w/w）", 2)):
            tk.Label(hdr, text=t, bg=COLOR_NAV, fg="#FFFFFF",
                     font=("Microsoft YaHei", 10, "bold"), anchor="w", width=w,
                     padx=10, pady=5).pack(side="left", fill="x", expand=True)
        self._comp_container = tk.Frame(self._form_inner, bg=COLOR_PANEL)
        self._comp_container.pack(fill="x", padx=14)
        self._comp_row_vars.clear()
        for comp in self.s3_components:
            self._add_comp_row(comp)
        b = tk.Button(self._form_inner, text="＋ 添加成分行", command=self._add_comp_row,
                      bg="#E8EAED", fg=COLOR_TEXT, relief="flat", padx=12, pady=4,
                      font=("Microsoft YaHei", 9), cursor="hand2")
        b.pack(anchor="w", padx=14, pady=(6, 2))

    def _add_comp_row(self, prefill=None):
        if self._comp_container is None:
            return
        prefill = prefill if prefill is not None else ("", "", "")
        row = tk.Frame(self._comp_container, bg=COLOR_PANEL)
        row.pack(fill="x", pady=2)
        vars_row = []
        for i in range(3):
            var = tk.StringVar(value=prefill[i])
            e = tk.Entry(row, textvariable=var, font=("Microsoft YaHei", 10), relief="solid",
                         highlightthickness=1, highlightbackground=COLOR_BORDER)
            e.pack(side="left", fill="x", expand=True, ipady=2, padx=4)
            vars_row.append(var)
        tk.Button(row, text="✕",
                  command=lambda r=row, v=vars_row: self._remove_comp_row(r, v),
                  bg="#FCE8E6", fg="#C5221F", relief="flat", padx=6, cursor="hand2",
                  font=("Microsoft YaHei", 9)).pack(side="left")
        self._comp_row_vars.append(vars_row)

    def _remove_comp_row(self, frame, vars_row):
        frame.destroy()
        try:
            self._comp_row_vars.remove(vars_row)
        except ValueError:
            pass

    def _collect_components(self):
        out = []
        for name, cas, conc in self.s3_components:
            if any((name, cas, conc)):
                out.append(ComponentRow(name, cas, conc))
        return out

    def _render_s9(self):
        self._section_header("S9 物理和化学特性",
                             "表格：序号 ｜ 标签（可编辑）｜ 输入框 —— 增删行自动重排序号 9.1~9.N")
        hdr = tk.Frame(self._form_inner, bg=COLOR_NAV)
        hdr.pack(fill="x", padx=14, pady=(8, 2))
        for t, expand in (("序号", False), ("标签", True), ("输入框", True), ("", False)):
            lkw = dict(font=("Microsoft YaHei", 10, "bold"), anchor="w", padx=6, pady=5)
            pkw = dict(side="left")
            if expand:
                pkw.update(fill="x", expand=True)
            else:
                lkw["width"] = 6
            tk.Label(hdr, text=t, bg=COLOR_NAV, fg="#FFFFFF", **lkw).pack(**pkw)
        self._s9_container = tk.Frame(self._form_inner, bg=COLOR_PANEL)
        self._s9_container.pack(fill="x", padx=14)
        self._s9_row_vars.clear()
        for i, data in enumerate(self.s9_data):
            self._add_s9_row(i, data["label"], data["value"])
        tk.Button(self._form_inner, text="＋ 添加行", command=self._add_new_s9_row,
                  bg="#E8EAED", fg=COLOR_TEXT, relief="flat", padx=12, pady=4,
                  font=("Microsoft YaHei", 9), cursor="hand2").pack(anchor="w", padx=14, pady=(6, 2))

    def _add_s9_row(self, idx, label="", value=""):
        """渲染一行 S9 表格行: 序号 | 标签(可编辑) | 输入框. idx 对应数据模型索引."""
        row = tk.Frame(self._s9_container, bg=COLOR_PANEL)
        row.pack(fill="x", pady=1)
        seq_label = tk.Label(row, text="9.%d" % (idx + 1), bg=COLOR_PANEL, fg="#80868B",
                             font=("Microsoft YaHei", 10), width=6, anchor="w", padx=6)
        seq_label.pack(side="left")
        vlabel = tk.StringVar(value=label)
        tk.Entry(row, textvariable=vlabel, font=("Microsoft YaHei", 10), relief="solid",
                 highlightthickness=1, highlightbackground=COLOR_BORDER).pack(
                 side="left", fill="x", expand=True, ipady=2, padx=4)
        vvalue = tk.StringVar(value=value)
        tk.Entry(row, textvariable=vvalue, font=("Microsoft YaHei", 10), relief="solid",
                 highlightthickness=1, highlightbackground=COLOR_BORDER).pack(
                 side="left", fill="x", expand=True, ipady=2, padx=4)
        info = {"seq_label": seq_label, "vlabel": vlabel, "vvalue": vvalue}
        tk.Button(row, text="✕",
                  command=lambda i=idx: self._remove_s9_row(i),
                  bg="#FCE8E6", fg="#C5221F", relief="flat", padx=6, cursor="hand2",
                  font=("Microsoft YaHei", 9)).pack(side="left", padx=(2, 6))
        self._s9_row_vars.append(info)

    def _add_new_s9_row(self):
        self._commit_current()
        self.s9_data.append({"label": "", "value": ""})
        self._render_current()

    def _remove_s9_row(self, idx):
        self._commit_current()
        if 0 <= idx < len(self.s9_data):
            del self.s9_data[idx]
        self._render_current()

    def _export(self):
        self._commit_current()
        if not self.s3_product_type.strip():
            messagebox.showwarning("缺少必填", "请选择产品类型（S3）")
            self._show_section("3")
            return
        s1 = self.s1_data
        comps = self._collect_components()
        s9_rows = [(row["label"], row["value"]) for row in self.s9_data]
        write_items = form_to_write_items(self._schema, s1, self.s3_product_type,
                                          comps, s9_rows)
        n_s1 = len(write_items["sections"]["1"])
        n_s9 = len(write_items["sections"]["9"])
        n_comp = len(write_items["sections"]["3"]["components"])
        default = "写入项_%s.json" % datetime.now().strftime("%Y%m%d_%H%M%S")
        path = filedialog.asksaveasfilename(
            title="生成写入项 JSON", defaultextension=".json", initialfile=default,
            filetypes=[("JSON", "*.json")])
        if not path:
            return
        Path(path).write_text(json.dumps(write_items, ensure_ascii=False, indent=2),
                              encoding="utf-8")
        messagebox.showinfo(
            "生成成功",
            "写入项已生成:\n%s\n\nS1: %d 字段 | S3: %d 成分 | S9: %d 字段\n\n"
            % (path, n_s1, n_comp, n_s9))


class FormWindow(tk.Toplevel):
    """独立表单窗口：内部直接复用 ``S139FormPanel``。"""

    def __init__(self, master):
        super().__init__(master)
        self.title("表单系统 · 录入 S1/S3/S9（参照 PEA-4139）")
        self.geometry("1120x760")
        self.configure(bg=COLOR_BG)
        self.minsize(980, 640)
        self.panel = S139FormPanel(self, show_actions=True, on_close=self.destroy)
        self.panel.pack(fill="both", expand=True)


__all__ = ["FormWindow", "S139FormPanel"]
