# -*- coding: utf-8 -*-
"""数据库 Tab: 一二级类目树 + 型号回读呈现 + 值编辑 + 导入检查 + 待确认池.

流程闭环: 导入MSDS → 检查报告 → 确认添加 → 写入总库 → 类目树回读 →
双击字段编辑 → 写库 → 重建字典(半自动标准字段) → 待确认池裁决.
"""
from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import ttk, filedialog, messagebox

from core.docx_reader import read_msds
from db import catalog, check, stddict
from db.schema import connect

from .section_tree import SectionView
from .theme import COLOR_NAV, COLOR_BORDER, COLOR_GRAY, COLOR_PANEL, COLOR_TEXT

_SELECTED = "selected"


class DbView(ttk.Frame):
    """数据库 Tab 主组件."""

    def __init__(self, master, db_path=None, on_status=None):
        super().__init__(master)
        self.on_status = on_status
        self.conn = None
        self.db_path = None
        self._result = None
        self._pid = None
        self._fid_by_key: dict[tuple, int] = {}
        self._build()
        if db_path:
            self.open_db(db_path)

    # ---------- 状态 ----------

    def _status(self, msg):
        if self.on_status:
            self.on_status(msg)

    # ---------- 布局 ----------

    def _build(self):
        bar = tk.Frame(self, bg="#EDF0F5")
        bar.pack(fill="x", padx=4, pady=(4, 4))
        ttk.Button(bar, text="新建库", command=self._new_db).pack(side="left")
        ttk.Button(bar, text="打开库", command=self._open_db).pack(side="left", padx=(4, 0))
        ttk.Button(bar, text="导入 MSDS", command=self._import_doc).pack(side="left", padx=(14, 0))
        ttk.Button(bar, text="重建字典", command=self._rebuild_dict).pack(side="left", padx=(4, 0))
        ttk.Button(bar, text="导出宽表", command=self._export_wide).pack(side="left", padx=(4, 0))
        self.db_label = tk.Label(bar, text="未打开数据库", bg="#EDF0F5",
                                 fg=COLOR_GRAY, font=("Microsoft YaHei", 9))
        self.db_label.pack(side="right", padx=6)

        body = ttk.PanedWindow(self, orient="horizontal")
        body.pack(fill="both", expand=True, padx=4, pady=(0, 4))
        body.add(self._build_left(), weight=0)

        nb = ttk.Notebook(body)
        self.tab_detail = ttk.Frame(nb)
        self.tab_pending = ttk.Frame(nb)
        nb.add(self.tab_detail, text="型号详情")
        nb.add(self.tab_pending, text="待确认池")
        self._build_detail_tab()
        self._build_pending_tab()
        body.add(nb, weight=1)

    def _build_left(self) -> tk.Frame:
        """左侧: 单目录 = 型号目录 (点型号 → 右侧 section).

        双目录原则: 左 = 型号, 右 = section. 不再有类目树/切换链接 (去掉三目录重叠).
        """
        left = tk.Frame(self, bg="#F5F6F8", width=300)
        # 标题行
        head = tk.Frame(left, bg="#F5F6F8")
        head.pack(fill="x", padx=8, pady=(6, 2))
        self._view_title = tk.Label(head, text="", bg="#F5F6F8",
                                    font=("Microsoft YaHei", 10, "bold"), fg=COLOR_TEXT)
        self._view_title.pack(side="left")

        # 检索框 (型号检索)
        self.search_var = tk.StringVar()
        e = tk.Entry(left, textvariable=self.search_var, font=("Microsoft YaHei", 10))
        e.pack(fill="x", padx=8)
        e.bind("<Return>", lambda _ev: self._search())
        ttk.Button(left, text="检索型号", command=self._search).pack(fill="x", padx=8, pady=3)

        # 型号目录: 以产品型号为索引平铺 (列: 型号 | 类目 | 语言 | 产品名)
        style = ttk.Style()
        style.configure("Cat.Treeview", rowheight=28, font=("Microsoft YaHei", 10))
        cols = ("model", "cat", "lang", "product")
        self.model_tree = ttk.Treeview(left, columns=cols, show="headings", selectmode="browse",
                                       style="Cat.Treeview")
        heads = {"model": "型号", "cat": "类目", "lang": "语言", "product": "产品名称"}
        widths = {"model": 92, "cat": 86, "lang": 34, "product": 80}
        for c in cols:
            self.model_tree.heading(c, text=heads[c])
            self.model_tree.column(c, width=widths[c], anchor="w",
                                   stretch=(c in ("model", "product")))
        self._model_vsb = ttk.Scrollbar(left, orient="vertical", command=self.model_tree.yview)
        self.model_tree.configure(yscrollcommand=self._model_vsb.set)
        self.model_tree.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=(2, 2))
        self._model_vsb.pack(side="right", fill="y", pady=(2, 2))
        self.model_tree.bind("<<TreeviewSelect>>", self._on_model_select)
        return left

    def _update_view_header(self):
        n = len(self.model_tree.get_children())
        self._view_title.configure(text=f"📋 型号目录 ({n})")

    def _load_model_tree(self, models=None):
        """型号目录: 平铺全部型号 (按型号排序), 支持过滤."""
        self.model_tree.delete(*self.model_tree.get_children())
        if not self.conn:
            return
        prods = catalog.list_products(self.conn)
        if models is not None:
            shown = set(models)
            prods = [p for p in prods if p["model_name"] in shown]
        prods.sort(key=lambda p: (p["model_name"].lower(), p.get("category_name") or ""))
        for p in prods:
            self.model_tree.insert("", "end", iid=f"m{p['id']}", values=(
                p["model_name"], p.get("category_name") or "",
                p.get("language") or p.get("template_vendor") or "",
                p.get("product_name") or ""),
                tags=(p["id"],))
        self._update_view_header()

    def _on_model_select(self, _ev):
        sel = self.model_tree.selection()
        if not sel:
            return
        try:
            self._show_product(int(self.model_tree.item(sel[0], "tags")[0]))
        except (ValueError, IndexError):
            pass

    def _build_detail_tab(self):
        """型号详情: 顶部型号信息栏 + 16节切换条 + 节内容 (双目录: 左型号 → 右节)."""
        info = tk.Frame(self.tab_detail, bg="#FFFFFF", height=44)
        info.pack(fill="x", padx=4, pady=4)
        info.pack_propagate(False)
        self.info_var = tk.StringVar(value="左侧选择型号以回读呈现")
        tk.Label(info, textvariable=self.info_var, bg="#FFFFFF", fg=COLOR_TEXT,
                 font=("Microsoft YaHei", 10, "bold"), anchor="w").pack(side="left", padx=10)
        ttk.Button(info, text="重新载入", command=self._reload_current).pack(side="right", padx=4, pady=6)
        ttk.Button(info, text="删除型号", command=self._delete_current).pack(side="right", padx=4, pady=6)

        # 16 节切换条 (代替原先的 16 节导航树, 不占额外一列 → 消灭三目录)
        bar = tk.Frame(self.tab_detail, bg="#EDF0F5")
        bar.pack(fill="x", padx=4)
        self._sec_btns: list[tk.Button] = []
        for n in range(1, 17):
            b = tk.Button(bar, text=str(n), width=3, relief="flat", bg="#FFFFFF",
                          fg=COLOR_TEXT, font=("Microsoft YaHei", 9), cursor="hand2",
                          command=lambda num=n: self._show_section(num))
            b.pack(side="left", padx=1, pady=3)
            self._sec_btns.append(b)
        tk.Label(bar, text="节", bg="#EDF0F5", fg=COLOR_GRAY,
                 font=("Microsoft YaHei", 9)).pack(side="left", padx=(4, 0))

        self.section_view = SectionView(self.tab_detail)
        self.section_view.pack(fill="both", expand=True, padx=4, pady=(4, 4))

    def _build_pending_tab(self):
        bar = tk.Frame(self.tab_pending, bg=COLOR_PANEL)
        bar.pack(fill="x", padx=4, pady=4)
        ttk.Button(bar, text="刷新", command=self._refresh_pending).pack(side="left")
        ttk.Button(bar, text="采纳建议 (并入标准字段)", command=self._adopt_suggestion).pack(side="left", padx=(6, 0))
        ttk.Button(bar, text="设为独立标准", command=self._set_active).pack(side="left", padx=(6, 0))
        self.pending_var = tk.StringVar(value="")
        tk.Label(bar, textvariable=self.pending_var, bg=COLOR_PANEL,
                 font=("Microsoft YaHei", 9), fg=COLOR_GRAY).pack(side="left", padx=10)
        cols = ("节", "待确认标签", "频次", "建议合并到")
        self.pending_tree = ttk.Treeview(self.tab_pending, columns=cols, show="headings",
                                         selectmode="browse")
        for c in cols:
            self.pending_tree.heading(c, text=c)
            self.pending_tree.column(c, width=90 if c in ("节", "频次") else 240, anchor="w")
        self.pending_tree.pack(fill="both", expand=True, padx=4, pady=(0, 4))
        self.pending_tree.bind("<<TreeviewSelect>>", lambda _e: None)
        tk.Label(self.tab_pending, text="提示: 选中一行, 采纳建议把标签并入相似标准字段 (保留原表达为别名); "
                                        "无建议时可选设为独立标准字段.",
                 bg=COLOR_PANEL, fg=COLOR_GRAY, font=("Microsoft YaHei", 9)).pack(fill="x", padx=6, pady=(0, 6))

    # ---------- 数据库开关 ----------

    def open_db(self, path: str):
        try:
            self.conn = connect(path)
            self.db_path = path
            self.db_label.configure(text=f"库: {Path(path).name}")
            self._refresh_pending()
            self._load_model_tree()
            self._status(f"已打开数据库 {Path(path).name}")
        except Exception as exc:
            messagebox.showerror("打开失败", str(exc))

    def _new_db(self):
        p = filedialog.asksaveasfilename(defaultextension=".db",
                                         filetypes=[("SQLite", "*.db")],
                                         initialfile="msds_total.db")
        if p:
            connect(p)
            self.open_db(p)

    def _open_db(self):
        p = filedialog.askopenfilename(filetypes=[("SQLite", "*.db"), ("所有文件", "*.*")])
        if p:
            self.open_db(p)

    # ---------- 型号检索 ----------

    def _search(self):
        if not self.conn:
            messagebox.showinfo("提示", "请先打开/新建数据库")
            return
        q = self.search_var.get().strip()
        if not q:
            self._load_model_tree()
            return
        hits = catalog.search_products(self.conn, q)
        self._load_model_tree(models={p["model_name"] for p in hits})
        self._status(f"检索「{q}」命中 {len(hits)} 个型号")

    # ---------- 型号呈现 / 编辑 ----------

    def _show_product(self, pid: int):
        self._pid = pid
        r = catalog.rebuild_product(self.conn, pid)
        if r is None:
            return
        self._result = r
        self._build_fid_map(pid)
        p = catalog.get_product(self.conn, pid)
        cat = (p or {}).get("category_name") or "未分类"
        fn = (p or {}).get("file_name") or ""
        self.info_var.set(f"🔍 {p['model_name']}  |  {cat}  |  修订 {(p or {}).get('revision_date') or '-'}  |  {fn}")
        self._highlight_section(1)
        self.section_view.show_section(1, r, on_value_edit=self._edit_value)

    def _show_section(self, num: int):
        if self._result is not None:
            self._highlight_section(num)
            self.section_view.show_section(num, self._result,
                                           on_value_edit=self._edit_value)

    def _highlight_section(self, num: int):
        """高亮当前选中的节按钮."""
        for i, b in enumerate(self._sec_btns, start=1):
            if i == num:
                b.configure(bg=COLOR_NAV, fg="#FFFFFF")
            else:
                b.configure(bg="#FFFFFF", fg=COLOR_TEXT)

    def _build_fid_map(self, pid: int):
        self._fid_by_key = {}
        groups: dict[int, list] = {}
        for f in catalog.get_fields(self.conn, pid):
            groups.setdefault(f["section"], []).append(f)
        for n, rows in groups.items():
            idx = 0
            for f in sorted(rows, key=lambda x: x["row_order"]):
                if f["kind"] == "section":
                    continue
                self._fid_by_key[(n, f["kind"], idx)] = f["id"]
                idx += 1

    def _edit_value(self, key: tuple[int, str, int], current: str):
        """双击值 → 弹窗编辑 → 写库."""
        fid = self._fid_by_key.get(key)
        if not fid:
            return
        num, kind, idx = key
        win = tk.Toplevel(self)
        win.title(f"编辑字段值 · 节{num} {kind}#{idx}")
        win.geometry("560x360")
        win.transient(self)
        tk.Label(win, text=f"节{num} · {kind} #{idx}", font=("Microsoft YaHei", 10),
                 anchor="w").pack(fill="x", padx=10, pady=(10, 2))
        box = tk.Text(win, font=("Microsoft YaHei", 10), wrap="word")
        box.insert("1.0", current or "")
        box.pack(fill="both", expand=True, padx=10)

        def save():
            catalog.update_field_value(self.conn, fid, box.get("1.0", "end-1c"))
            win.destroy()
            self._reload_current()
            self._status(f"已保存 节{num} · {kind}#{idx}")

        bar = tk.Frame(win)
        bar.pack(fill="x", padx=10, pady=8)
        ttk.Button(bar, text="保存", command=save).pack(side="left")
        ttk.Button(bar, text="取消", command=win.destroy).pack(side="left", padx=(6, 0))

    def _reload_current(self):
        if self._pid:
            self._show_product(self._pid)

    def _delete_current(self):
        if not self._pid:
            return
        if not messagebox.askyesno("确认删除",
                                   "软删除该型号? (active=0, 可从库恢复/不再展示)"):
            return
        catalog.delete_product(self.conn, self._pid)
        self._pid = None
        self._result = None
        self.info_var.set("已软删除该型号")
        self._load_model_tree()
        self._status("已软删除型号")

    # ---------- 导入流程 ----------

    def _import_doc(self):
        if not self.conn:
            messagebox.showinfo("提示", "请先打开/新建数据库")
            return
        path = filedialog.askopenfilename(filetypes=[("Word 文档", "*.docx"),
                                                     ("所有文件", "*.*")])
        if not path:
            return
        try:
            result = read_msds(Path(path))
        except Exception as exc:
            messagebox.showerror("读取失败", str(exc))
            return
        ImportDialog(self, result, self.conn, on_import=self._after_import)

    def _after_import(self):
        self._load_model_tree()
        self._refresh_pending()
        self._status("导入完成，型号目录与待确认池已刷新")

    # ---------- 字典 / 导出 ----------

    def _rebuild_dict(self):
        if not self.conn:
            return
        n = catalog.db_stats(self.conn)["fields"]
        if not messagebox.askyesno("重建字典",
                                   f"将从库内 {n} 条字段行重新扫描并重建标准字段字典.\n"
                                   "已有 active 条目作为同义组种子保留; 重建会覆盖 pending 列表.\n继续?"):
            return
        try:
            stats = stddict.scan_db(self.conn)
            res = stddict.build_field_dict(self.conn, stats=stats)
            self._refresh_pending()
            self._status(f"字典重建完成: 标准组 {res.get('groups', 0)} / 待确认 {res.get('pending', 0)} / "
                         f"统计标签 {res.get('stats_count', 0)}")
        except Exception as exc:
            messagebox.showerror("重建失败", str(exc))

    def _export_wide(self):
        if not self.conn:
            return
        p = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                         filetypes=[("Excel", "*.xlsx")],
                                         initialfile="MSDS总库_宽表.xlsx")
        if not p:
            return
        try:
            from db.export import export_wide
            export_wide(self.conn, p)
            self._status(f"宽表已导出: {Path(p).name}")
            messagebox.showinfo("完成", f"宽表已导出到:\n{p}")
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc))

    # ---------- 待确认池 ----------

    def _refresh_pending(self):
        for i in self.pending_tree.get_children():
            self.pending_tree.delete(i)
        if not self.conn:
            return
        rows = stddict.list_pending(self.conn)
        sugg = stddict.suggest_merges(self.conn, top_k=1, min_sim=0.5)
        by = {(s["section"], s["pending"]): s["candidates"][0] for s in sugg}
        for p in rows:
            s = by.get((p["section"], p["std_field"]))
            cand = f"{s['target']} (相似{s['similarity']})" if s else ""
            target = s["target"] if s else ""
            self.pending_tree.insert("", "end",
                                     values=(p["section"], p["std_field"], p["freq"], cand),
                                     tags=(target,))
        self.pending_var.set(f"待确认 {len(rows)} 项 · 双击操作按钮处理")

    def _selected_pending(self):
        sel = self.pending_tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先在列表中选中一行")
            return None
        v = self.pending_tree.item(sel[0], "values")
        sec = int(v[0]); std = v[1]
        tags = self.pending_tree.item(sel[0], "tags")
        target = tags[0] if tags else ""
        return sec, std, target

    def _adopt_suggestion(self):
        if not self.conn:
            return
        sp = self._selected_pending()
        if not sp:
            return
        sec, std, target = sp
        if not target:
            messagebox.showinfo("提示", "该标签没有可采纳的建议, 可改为「设为独立标准」")
            return
        if stddict.merge_pending(self.conn, std, sec, target):
            self._refresh_pending()
            self._status(f"已采纳: {std} → {target} (原表达保留为别名)")
        else:
            messagebox.showerror("失败", "目标标准字段不存在, 请先刷新")

    def _set_active(self):
        if not self.conn:
            return
        sp = self._selected_pending()
        if not sp:
            return
        sec, std, _ = sp
        stddict.set_pending_active(self.conn, std, sec)
        self._refresh_pending()
        self._status(f"已设为独立标准字段: {std}")


class ImportDialog(tk.Toplevel):
    """导入检查报告: 缺失节/必填/异常/未归并/重复 → 类目+型号 → 添加到总库."""

    def __init__(self, master, result, conn, on_import=None):
        super().__init__(master)
        self.result = result
        self.conn = conn
        self.on_import = on_import
        self.title("导入检查报告")
        self.geometry("680x560")
        self.transient(master)
        self._build()

    def _build(self):
        tk.Label(self, text=f"文件: {Path(self.result.file_name).name}",
                 font=("Microsoft YaHei", 10, "bold"), anchor="w").pack(fill="x", padx=10, pady=(10, 2))
        cols = ("类别", "内容")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=12)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=90 if c == "类别" else 540, anchor="w")
        self.tree.pack(fill="both", expand=True, padx=10, pady=4)

        rpt = check.check_doc(self.result, self.conn)
        self._row("型号", rpt["model"])
        self._row("节数/字段/成分/象形图",
                  f"{16 - len(rpt['missing_sections'])}节 · {rpt['fields_count']}字段 · "
                  f"{rpt['components_count']}成分 · {rpt['pictograms_count']}象形图")
        if rpt["missing_sections"]:
            self._row("缺失节", "、".join(f"S{n}" for n in rpt["missing_sections"]))
        if rpt["required_missing"]:
            self._row("必填缺失", "、".join(rpt["required_missing"]))
        for a in rpt["anomalies"]:
            self._row(f"异常(S{a['section']})", a["message"])
        if rpt["duplicate_sha"]:
            self._row("重复", f"sha256 与型号 {rpt['duplicate_sha']['model']} 相同 (将跳过写入)")
        if rpt["duplicate_model"]:
            self._row("重复", f"同类目下型号 {rpt['duplicate_model']['model']} 已存在 (将跳过写入)")
        if rpt["unmatched_labels"]:
            labs = "、".join(f"[S{u['section']}]{u['label']}" for u in rpt["unmatched_labels"][:8])
            more = f" 等{len(rpt['unmatched_labels'])}项" if len(rpt["unmatched_labels"]) > 8 else ""
            self._row("未归并标签", labs + more + " (入库时保留原始表达, 列归待确认池)")

        foot = tk.Frame(self)
        foot.pack(fill="x", padx=10, pady=(4, 10))
        tk.Label(foot, text="一级类目:", font=("Microsoft YaHei", 10)).pack(side="left")
        self.cat_var = tk.StringVar()
        cats = [c["name"] for c in catalog.list_categories(self.conn)]
        self.cat_cb = ttk.Combobox(foot, textvariable=self.cat_var, values=cats, width=16)
        self.cat_cb.pack(side="left", padx=(4, 12))
        tk.Label(foot, text="型号:", font=("Microsoft YaHei", 10)).pack(side="left")
        self.model_var = tk.StringVar(value=rpt["model"])
        ttk.Entry(foot, textvariable=self.model_var, width=20).pack(side="left", padx=(4, 12))
        ttk.Button(foot, text="添加到总库", command=self._confirm).pack(side="right")

    def _row(self, cat, content):
        self.tree.insert("", "end", values=(cat, content))

    def _confirm(self):
        cat = self.cat_var.get().strip()
        model = self.model_var.get().strip()
        if not model:
            messagebox.showwarning("提示", "型号不能为空")
            return
        try:
            pid, status = catalog.add_product(self.conn, self.result,
                                              category=cat, model_name=model)
        except Exception as exc:
            messagebox.showerror("写入失败", str(exc))
            return
        self.destroy()
        if status == "exists_sha":
            messagebox.showinfo("重复", f"sha256 相同, 未重复添加 (已有 id={pid})")
        elif status == "exists_model":
            messagebox.showinfo("重复", f"同类目下型号 {model} 已存在, 未重复添加")
        else:
            messagebox.showinfo("成功", f"已写入: {cat or '(未分类)'} / {model}")
        if self.on_import:
            self.on_import()
