# -*- coding: utf-8 -*-
"""16 节结构化显示: 树形 + 与 Word 模板表格对齐的节内容视图 (含覆写状态徽章)."""
from __future__ import annotations

import re
import tkinter as tk
from tkinter import ttk

from core.extract import build_hierarchy, search_tree
from core.structure import ParseResult

from .theme import (
    COLOR_BORDER, COLOR_GRAY, COLOR_GREEN, COLOR_NAV, COLOR_NAV_SEL,
    COLOR_PANEL, COLOR_ROW_ALT, COLOR_TEXT,
)


class SectionTree(ttk.Frame):
    """左侧 16 节导航树."""

    def __init__(self, master, on_select=None):
        super().__init__(master, width=230)
        self.on_select = on_select
        self._build()

    def _build(self):
        style = ttk.Style()
        style.configure("Nav.Treeview", background=COLOR_NAV, foreground="#E8EAED",
                        fieldbackground=COLOR_NAV, rowheight=30, borderwidth=0,
                        font=("Microsoft YaHei", 10))
        style.configure("Nav.Treeview.Heading", background=COLOR_NAV, foreground="#E8EAED",
                        borderwidth=0)
        style.map("Nav.Treeview", background=[("selected", COLOR_NAV_SEL)],
                  foreground=[("selected", "#FFFFFF")])

        self.tree = ttk.Treeview(self, style="Nav.Treeview", show="tree", selectmode="browse")
        self.tree.heading("#0", text="MSDS 16 节")
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.pack(fill="both", expand=True)
        self._items: dict[int, str] = {}

    def set_result(self, result: ParseResult):
        """加载完整三级树 (保存原始节点, 供检索过滤后恢复)."""
        self._result = result
        self._full_nodes = build_hierarchy(result)
        self._apply_nodes(self._full_nodes, open_sec=True)

    def filter_by(self, query: str, scope: str = "all"):
        """按关键词过滤导航树: 命中节/字段保留, 其余隐藏.

        保留三级父子关系 (命中字段连带其节/大标题一起显示).
        query 为空 → 恢复完整树.
        """
        if not self._result:
            return
        q = query.strip()
        if not q:
            self._apply_nodes(self._full_nodes, open_sec=True)
            return
        nodes = search_tree(self._full_nodes, q, scope)
        self._apply_nodes(nodes, open_sec=True)

    def _apply_nodes(self, nodes, open_sec: bool = True):
        """用给定节点重建导航树 (只显示节一级, 节号映射随之更新)."""
        self.tree.delete(*self.tree.get_children())
        self._items.clear()
        for sn in nodes:
            iid = self.tree.insert("", "end", text=f"{sn.number}. {sn.title}",
                                   open=open_sec)
            self._items[sn.number] = iid

    def select(self, num: int):
        iid = self._items.get(num)
        if iid:
            self.tree.selection_set(iid)
            self.tree.see(iid)

    def _on_select(self, _event):
        """目录只放节: 选中任一节点直接映射到所属节号."""
        if self.on_select:
            sel = self.tree.selection()
            if not sel:
                return
            iid = sel[0]
            for n, item in self._items.items():
                if item == iid:
                    self.on_select(n)
                    break


class SectionView(ttk.Frame):
    """右上: 当前节内容, 按 Word 模板表格方式渲染 (标题|内容 两列).

    布局与模板表格对齐:
      - 节标题行 (全宽, 深蓝, 加粗)      ← 模板表格 R0
      - 字段行 (徽章|标签|值, 值自动换行)   ← 模板表格 标签|值 行
      - 子标题/说明行 也拆 标题|内容 两列    ← 模板表格 单列通栏
      - 成分表 (S3, 三列带表头)          ← 模板表格 S3 成分表
    每行带 可编辑/不可编辑 徽章, 点击徽章手动标注字段权限:
      - 可编辑 = 后续可被覆写替换
      - 不可编辑 = 固定字段
    """

    def __init__(self, master):
        super().__init__(master, padding=6)
        self._overrides: dict[tuple[int, str, int], bool] = {}
        self._on_toggle = None
        self._on_value_edit = None   # 可选: 双击值 → 编辑回调 (key, 当前值)
        self._num = 1
        self._result: ParseResult | None = None
        self._value_labels: list[tk.Label] = []
        self._label_labels: list[tk.Label] = []
        self._span_labels: list[tk.Label] = []   # 跨列/通栏行内容 Label (自动换行)
        self._photos: list = []          # 象形图 PhotoImage 引用 (防 GC 清空)
        self._img_index = 0              # 该节图片分配游标
        self._ri = 0
        self._table: tk.Frame | None = None
        self._table_ri = 0
        self._build()

    def _build(self):
        self.title_var = tk.StringVar()
        tk.Label(self, textvariable=self.title_var, font=("Microsoft YaHei", 13, "bold"),
                 bg=COLOR_PANEL, fg=COLOR_TEXT, anchor="w").pack(fill="x", pady=(0, 2))

        # 页眉/页脚摘要条
        self.meta_var = tk.StringVar()
        tk.Label(self, textvariable=self.meta_var, font=("Microsoft YaHei", 9),
                 bg=COLOR_PANEL, fg=COLOR_GRAY, anchor="w").pack(fill="x", pady=(0, 2))

        # 字段权限图例
        self.legend_var = tk.StringVar()
        tk.Label(self, textvariable=self.legend_var, font=("Microsoft YaHei", 9),
                 bg=COLOR_PANEL, fg=COLOR_GRAY, anchor="w").pack(fill="x", pady=(0, 4))

        # 可滚动画布
        wrap = tk.Frame(self, bg=COLOR_PANEL)
        wrap.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(wrap, bg=COLOR_PANEL, highlightthickness=0)
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.inner = tk.Frame(self.canvas, bg=COLOR_PANEL)
        self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", self._on_inner_cfg)
        self.canvas.bind("<Configure>", self._on_canvas_cfg)
        self.canvas.bind("<MouseWheel>", lambda e: self.canvas.yview_scroll(int(-e.delta / 120), "units"))

    # ---------- 画布/滚动 ----------

    def _on_inner_cfg(self, _e):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_cfg(self, e):
        self.canvas.itemconfigure(self._win, width=e.width)
        # 固定列: 徽章(~90) + 序号(~70) + 内外框边距(~40)
        avail = max(320, e.width - 200)
        # 标题列: 弹性给 ~35%, 内容列: ~60% (余量内框), 均有下限 → 长标题自动换行不截断
        wrap_lbl = max(90, int(avail * 0.35))
        wrap_val = max(140, avail - wrap_lbl - 20)
        for l in self._label_labels:
            l.configure(wraplength=wrap_lbl)
        for v in self._value_labels:
            v.configure(wraplength=wrap_val)
        # 跨列/通栏行: 内容几乎占满整表宽 (徽章 + 内框边距除外)
        wrap_span = max(160, avail - 40)
        for v in self._span_labels:
            v.configure(wraplength=wrap_span)

    # ---------- 数据渲染 ----------

    def show_section(self, num: int, result: ParseResult,
                     overrides: dict[tuple[int, str, int], bool] | None = None,
                     on_toggle=None, on_value_edit=None):
        """渲染一节. overrides: 手动标注的编辑状态; on_toggle: 切换回调;
        on_value_edit: 双击值单元格回调 (key, 当前值), 供数据库编辑."""
        self._num = num
        self._result = result
        if overrides is not None:
            self._overrides = overrides
        self._on_toggle = on_toggle
        self._on_value_edit = on_value_edit

        for w in self.inner.winfo_children():
            w.destroy()
        self._ri = 0
        self._value_labels = []
        self._label_labels = []
        self._span_labels = []
        self._photos = []
        self._img_index = 0

        sec = result.sections.get(num)
        if not sec:
            self.title_var.set(f"第{num}节 (缺失)")
            self.meta_var.set("")
            self.legend_var.set("")
            return
        self.title_var.set(f"第{num}节 · {sec.full_title}")

        if num == 0:
            self.meta_var.set("此节为页眉/页脚字段, 独立管理 可编辑/不可编辑 (不参与自动覆写比对)")
        else:
            head = result.header.replace("\n", " | ") if result.header else "无页眉"
            foot = result.footer.replace("\n", " | ") if result.footer else "无页脚"
            self.meta_var.set(f"页眉: {head}    页脚: {foot}")

        self.legend_var.set("三列表格: 序号🔒标题列固定不可覆写 · 🟢内容可编辑 ⚪内容不可编辑 — 点击徽章切换内容列")

        # 整节渲染为一张连续表格: 外框(深蓝2px) + 内框(浅灰1px 行列分隔线)
        self._begin_table()
        for row in sec.iter_rows():
            self._render_row(row)

        # 成分表 (S3)
        if sec.is_component_table:
            self._component_table(sec)

        if not sec.fields and not sec.lines and not sec.is_component_table:
            self._title_row("(本节无字段 / 说明行 / 成分)")

        # 按画布实际宽度刷新 标题/内容 列换行 (长标题换行显示全, 自适应)
        w = self.canvas.winfo_width()
        if w > 1:
            self._on_canvas_cfg(type("E", (), {"width": w}))

    # ---------- 行渲染 ----------

    def _editable_of(self, key: tuple[int, str, int], default: bool) -> bool:
        return self._overrides.get(key, default)

    # ---------- 表格容器 (外框 + 内框线) ----------

    def _begin_table(self):
        """创建整节表格: 外框 2px 深蓝 + 内框线底色 (浅灰)."""
        outer = tk.Frame(self.inner, bg=COLOR_BORDER, highlightthickness=2,
                         highlightbackground=COLOR_NAV)
        outer.grid(row=self._ri, column=0, sticky="ew")
        self._ri += 1
        self._table = tk.Frame(outer, bg=COLOR_BORDER, highlightthickness=0)
        self._table.pack(fill="both", expand=True, padx=0, pady=0)
        self._table_ri = 0
        self._table.grid_columnconfigure(0, weight=0)   # 徽章
        self._table.grid_columnconfigure(1, weight=0)   # 序号
        self._table.grid_columnconfigure(2, weight=0)   # 标题
        self._table.grid_columnconfigure(3, weight=1)   # 内容 (弹性)

    def _cell(self, col: int, row_bg: str, factory) -> object:
        """建一个单元格 Frame (右下 1px 内框线), factory(cell) 放置内容并返回."""
        cell = tk.Frame(self._table, bg=row_bg)
        cell.grid(row=self._table_ri, column=col, sticky="nsew",
                  padx=(0, 0 if col == 3 else 1), pady=(0, 1))
        cell.grid_rowconfigure(0, weight=1)
        return factory(cell)

    def _text_cell(self, col, text, row_bg, fg, wrap=None, width=None, bold=False,
                   anchor="nw", padx=6, pady=7, font_size=10):
        """文本单元格: 文本高度/宽度随内容与容器自适应."""
        def factory(cell):
            w = tk.Label(cell, text=text, bg=row_bg, fg=fg,
                         font=("Microsoft YaHei", font_size, "bold" if bold else "normal"),
                         anchor=anchor, justify="left", width=width, wraplength=wrap)
            w.grid(row=0, column=0, sticky="nsew", padx=padx, pady=pady)
            return w
        return self._cell(col, row_bg, factory)

    def _badge_cell(self, key, editable: bool, row_bg: str, font_size=8):
        """徽章按钮单元格 (顶部对齐, 不占满行高)."""
        tag_text = "可编辑" if editable else "不可编辑"
        tag_color = COLOR_GREEN if editable else COLOR_GRAY
        def factory(cell):
            b = tk.Button(cell, text=tag_text, bg=tag_color, fg="#FFFFFF",
                          font=("Microsoft YaHei", font_size, "bold"), padx=4, pady=1,
                          relief="flat", cursor="hand2", width=6,
                          command=lambda k=key: self._toggle(k))
            b.grid(row=0, column=0, sticky="n", padx=8, pady=7)
            return b
        return self._cell(0, row_bg, factory)

    # ---------- 行渲染 ----------

    def _render_row(self, row):
        if row.kind == "section":
            self._title_row(row.label)
            return
        key = (self._num, row.kind, row.index)
        editable = self._editable_of(key, row.editable)
        if row.span:
            # 跨列/通栏行 (总结句等): 内容跨 序号|标题|内容 三列显示, 与 Word 一致
            self._span_row(key, row.value, editable)
        else:
            self._three_col_row(key, row.seq, row.label, row.value, editable)

    def _span_row(self, key, value, editable: bool):
        """跨列/通栏行: 徽章 | 内容(跨 序号+标题+内容 三列).

        识别 Word 中跨列合并的总结句/通栏说明 (如 S11 '该产品无可用的毒理学研究。'),
        内容不再挤在单独内容列, 而是像原文档一样通栏显示.
        """
        bg = COLOR_ROW_ALT if (self._table_ri % 2 == 0) else COLOR_PANEL
        self._badge_cell(key, editable, bg)
        # 内容跨 列1~3 (序号|标题|内容), 宽度自适应
        cell = tk.Frame(self._table, bg=bg)
        cell.grid(row=self._table_ri, column=1, columnspan=3, sticky="nsew",
                  padx=(0, 0), pady=(0, 1))
        cell.grid_rowconfigure(0, weight=1)
        cell.grid_columnconfigure(0, weight=1)
        fg = COLOR_TEXT if editable else COLOR_GRAY
        v = tk.Label(cell, text=value, bg=bg, fg=fg,
                     font=("Microsoft YaHei", 10), anchor="nw", justify="left",
                     wraplength=560)
        v.grid(row=0, column=0, sticky="nsew", padx=10, pady=7)
        self._span_labels.append(v)
        self._table_ri += 1

    def _three_col_row(self, key, seq, label, value, editable: bool):
        """徽章 | 序号 | 标题(锁定) | 内容 四列表格行 (连续表格内).

        标题列自适应: wraplength 随窗口宽度重算, 长标题自动换行显示全, 不截断.
        内容列含 [象形图] 占位 → 显示该节对应的原图 (替换占位文本).
        """
        bg = COLOR_ROW_ALT if (self._table_ri % 2 == 0) else COLOR_PANEL
        self._badge_cell(key, editable, bg)
        self._text_cell(1, seq or "", bg, COLOR_TEXT, bold=True, width=8, padx=2)
        l = self._text_cell(2, f"🔒 {label}" if label else "", bg, COLOR_TEXT,
                            bold=True, wrap=260, padx=2)
        self._label_labels.append(l)
        val_fg = COLOR_TEXT if editable else COLOR_GRAY
        if "[象形图" in (value or ""):
            v = self._value_cell_with_images(value, bg, val_fg, wrap=560)
            if v is not None:
                self._value_labels.append(v)   # 文本 Label 参与 wrap 自适应
        else:
            v = self._text_cell(3, value, bg, val_fg, wrap=560, padx=6)
            self._value_labels.append(v)
        # 启用编辑回调时: 双击值单元格 → 编辑
        if self._on_value_edit and v is not None:
            v.bind("<Double-Button-1>",
                   lambda e, k=key, cur=value: self._edit_value(k, cur))
        self._table_ri += 1

    def _value_cell_with_images(self, value, row_bg, fg, wrap=None):
        """内容列单元格: 文本 + 该字段对应的象形图原图 (占位符替换为图).

        返回文本 Label (供 wrap 自适应); 无文本时返回 None.
        """
        n = 0
        m = re.search(r"\[象形图(?:×(\d+))?\]", value or "")
        if m:
            n = int(m.group(1) or 1)
        # 从该节 images 顺序取图 (占位符在 value 中出现的顺序)
        sec_imgs = [im for im in (self._result.images if self._result else [])
                    if im.section == self._num]
        take = sec_imgs[self._img_index:self._img_index + n]
        self._img_index += n
        text = re.sub(r"\[象形图(?:×\d+)?\]", "", value or "").strip()

        def factory(cell):
            # 文本 Label (可折行) + 象形图 Label (横向排)
            wrap_label = tk.Frame(cell, bg=row_bg)
            wrap_label.grid(row=0, column=0, sticky="nsew")
            tl = None
            if text:
                tl = tk.Label(wrap_label, text=text, bg=row_bg, fg=fg,
                              font=("Microsoft YaHei", 10), anchor="nw",
                              justify="left", wraplength=wrap)
                tl.grid(row=0, column=0, sticky="nw", padx=6, pady=7)
            col = 1 if text else 0
            for im in take:
                ph = self._photo_of(im.blob, im.ext)
                if ph is None:
                    continue
                il = tk.Label(wrap_label, image=ph, bg=row_bg,
                              cursor="hand2")
                il.image = ph
                il.grid(row=0, column=col, sticky="nw", padx=(6 if col else 6, 2),
                        pady=7)
                il.bind("<Button-1>",
                        lambda e, b=im.blob, x=im.ext: self._save_image(b, x))
                col += 1
            return tl if tl is not None else wrap_label
        return self._cell(3, row_bg, factory)

    def _photo_of(self, blob: bytes, ext: str):
        """blob → tk.PhotoImage (PNG 原图 / JPEG 经 PIL 转); 保持引用防 GC."""
        import io
        try:
            if ext == "jpeg" or ext == "jpg":
                from PIL import Image, ImageTk
                im = Image.open(io.BytesIO(blob)).convert("RGBA")
                im.thumbnail((160, 96), Image.LANCZOS)
                ph = ImageTk.PhotoImage(im)
            else:
                ph = tk.PhotoImage(data=blob)
        except Exception:
            return None
        self._photos.append(ph)
        return ph

    def _save_image(self, blob: bytes, ext: str) -> None:
        """点击象形图 → 保存原图到 结构读取/outputs/ 目录."""
        try:
            from tkinter import filedialog
            p = filedialog.asksaveasfilename(
                defaultextension=f".{ext}",
                filetypes=[(f"{ext.upper()} 图片", f"*.{ext}"),
                           ("所有文件", "*.*")],
                initialfile=f"象形图_s{self._num}.{ext}")
            if p:
                with open(p, "wb") as f:
                    f.write(blob)
        except Exception as exc:
            pass

    def _title_row(self, text: str):
        """通栏节标题行 (表头, 深蓝, 跨全宽)."""
        cell = tk.Frame(self._table, bg=COLOR_NAV)
        cell.grid(row=self._table_ri, column=0, columnspan=4, sticky="nsew",
                  padx=(0, 0), pady=(0, 1))
        tk.Label(cell, text=text, bg=COLOR_NAV, fg="#FFFFFF",
                 font=("Microsoft YaHei", 11, "bold"), anchor="w",
                 padx=12, pady=7).pack(fill="x")
        self._table_ri += 1

    def _component_table(self, sec):
        """S3 成分表: 表头 + 数据行, 融入连续表格 (统一内外框).

        与主表格一致: 每列独立 cell (列间 1px 内框竖线 + 行间横线),
        表头深蓝底, 徽章只控制成分行可编辑性.
        """
        # 表头: 优先用实际识别到的表头 (英文文件如 'Chemical Name | CAS Number | %（w/w）')
        hdr_parts = [p.strip() for p in (sec.component_header or "").split(" | ")]
        while len(hdr_parts) < 3:
            hdr_parts.append("")
        heads = (("", 6),
                 (hdr_parts[0] or "化学品名称", 18),
                 (hdr_parts[1] or "CAS编号", 15),
                 (hdr_parts[2] or "含量%（w/w）", 14))
        # 表头: 每列一个深蓝 cell
        for ci, (h, _w) in enumerate(heads):
            cell = tk.Frame(self._table, bg=COLOR_NAV)
            cell.grid(row=self._table_ri, column=ci, sticky="nsew",
                      padx=(0, 0 if ci == len(heads) - 1 else 1), pady=(0, 1))
            cell.grid_rowconfigure(0, weight=1)
            cell.grid_columnconfigure(0, weight=1)
            tk.Label(cell, text=h, bg=COLOR_NAV, fg="#FFFFFF",
                     font=("Microsoft YaHei", 10, "bold"), anchor="w",
                     padx=10, pady=5).grid(row=0, column=0, sticky="nsew")
        self._table_ri += 1

        for ci, comp in enumerate(sec.components):
            key = (self._num, "component", ci)
            editable = self._editable_of(key, comp.editable)
            bg = COLOR_ROW_ALT if (self._table_ri % 2 == 0) else COLOR_PANEL
            self._badge_cell(key, editable, bg, font_size=7)
            fg = COLOR_TEXT if editable else COLOR_GRAY
            for ci2, v in enumerate((comp.name, comp.cas, comp.conc)):
                self._text_cell(ci2 + 1, v, bg, fg, wrap=260,
                                width=heads[ci2 + 1][1], padx=8, pady=5)
            self._table_ri += 1

    # ---------- 值编辑 (数据库模式) ----------

    def _edit_value(self, key: tuple[int, str, int], current: str):
        """双击值单元格 → 回调上层 (由上层弹编辑框处理)."""
        if self._on_value_edit:
            self._on_value_edit(key, current)

    # ---------- 手动标注 ----------

    def _toggle(self, key: tuple[int, str, int]):
        """点击徽章: 切换该字段 可编辑/不可编辑."""
        cur = self._overrides.get(key)
        default = self._default_editable(key)
        new = not (cur if cur is not None else default)
        self._overrides[key] = new
        if self._on_toggle:
            self._on_toggle(key, new)
        if self._result is not None:
            self.show_section(self._num, self._result, self._overrides, self._on_toggle)

    def _default_editable(self, key: tuple[int, str, int]) -> bool:
        """未手动标注时的默认编辑状态 (来自数据模型)."""
        num, kind, idx = key
        sec = self._result.sections.get(num) if self._result else None
        if not sec:
            return True
        if kind == "component":
            return sec.components[idx].editable if idx < len(sec.components) else True
        for row in sec.iter_rows():
            if row.kind == kind and row.index == idx:
                return row.editable
        return True
