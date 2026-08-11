# -*- coding: utf-8 -*-
"""正式筛选库三表浏览: 中文库 / 英文库 / 宽表库 父子级目录.

父 = 三张表 (zh_products / en_products / wide_rows)
子 = 型号
孙 = 标准字段(统一库) 或 原始表述记录(宽表库)

右侧详情保留原 Word 连续表格形式 (见 screened_detail.TableSheet).
数据源: F:/正式项目与模块化内容/Word 覆写模块/数据库/正式筛选库/正式筛选库.db
原 msds_total.db 总库保留不动, 本视图是新的替换入口.
"""
from __future__ import annotations

import sqlite3
import tkinter as tk
from pathlib import Path
from tkinter import ttk, filedialog, messagebox

from .screened_detail import ScreenedDetail
from .theme import COLOR_BORDER, COLOR_BG, COLOR_GRAY, COLOR_PANEL, COLOR_TEXT

DEFAULT_DB = Path(r"F:\正式项目与模块化内容\Word 覆写模块\数据库\正式筛选库\正式筛选库.db")

_META_COLS = ("model", "vendors", "files", "created_at")  # 非字段列


class ScreenedView(ttk.Frame):
    """正式筛选库 三表父子级目录浏览."""

    def __init__(self, master, db_path=None, on_status=None):
        super().__init__(master)
        self.on_status = on_status
        self.conn: sqlite3.Connection | None = None
        self.db_path: Path | None = None
        self._zh_cols: list[str] = []
        self._en_cols: list[str] = []
        self._loaded_models: set = set()
        self._build()
        if db_path:
            self.open_db(db_path)

    # ---------- 状态 ----------

    def _status(self, msg):
        if self.on_status:
            self.on_status(msg)

    # ---------- 布局 ----------

    def _build(self):
        bar = tk.Frame(self, bg=COLOR_PANEL, highlightthickness=1, highlightbackground=COLOR_BORDER)
        bar.pack(fill="x", padx=4, pady=(4, 4))
        ttk.Button(bar, text="打开库", command=self._open_db).pack(side="left")
        ttk.Button(bar, text="刷新", command=self._refresh).pack(side="left", padx=(4, 0))
        ttk.Button(bar, text="导出当前", command=self._export_current).pack(side="left", padx=(4, 0))
        self.search_var = tk.StringVar()
        e = tk.Entry(bar, textvariable=self.search_var, font=("Microsoft YaHei", 10), width=22)
        e.pack(side="left", padx=(14, 0))
        e.bind("<Return>", lambda _ev: self._apply_search())
        ttk.Button(bar, text="检索", command=self._apply_search).pack(side="left", padx=(2, 0))
        self.db_label = tk.Label(bar, text="未打开库", bg=COLOR_PANEL,
                                 fg=COLOR_GRAY, font=("Microsoft YaHei", 9))
        self.db_label.pack(side="right", padx=6)

        body = tk.Frame(self, bg=COLOR_BG)
        body.pack(fill="both", expand=True, padx=4, pady=(0, 4))

        # 左侧: 三表父子级目录树
        self.tree = ttk.Treeview(body, selectmode="browse")
        style = ttk.Style()
        style.configure("S.Treeview", rowheight=26, font=("Microsoft YaHei", 10))
        style.configure("S.Treeview.Heading", font=("Microsoft YaHei", 10, "bold"))
        self.tree.configure(style="S.Treeview")
        vsb = ttk.Scrollbar(body, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="y", padx=(0, 4))
        vsb.pack(side="left", fill="y")
        self.tree.bind("<<TreeviewOpen>>", self._on_open)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        # 右侧: 详情区 (保留原 Word 连续表格形式)
        self.detail = ScreenedDetail(self, on_status=self._status)
        self.detail.pack(side="left", fill="both", expand=True)

    # ---------- 库打开 ----------

    def open_db(self, path):
        try:
            if self.conn:
                self.conn.close()
            self.conn = sqlite3.connect(str(path))
            self.conn.row_factory = sqlite3.Row
            self.db_path = Path(path)
        except Exception as exc:
            messagebox.showerror("打开失败", str(exc))
            return
        self._zh_cols = self._std_cols("zh_products")
        self._en_cols = self._std_cols("en_products")
        self._loaded_models.clear()
        self.tree.delete(*self.tree.get_children())
        for iid, label in (("zh", "中文库"), ("en", "英文库"), ("wide", "宽表库")):
            self.tree.insert("", "end", iid=iid, text=label, open=False)
        self._refresh_counts()
        # 默认展开中文库, 懒加载型号 (首屏即有内容)
        self.tree.item("zh", open=True)
        self._load_models("zh", "zh_products")
        self._status("已打开正式筛选库: %s" % self.db_path.name)

    def _std_cols(self, table):
        cols = [r[1] for r in self.conn.execute('PRAGMA table_info("%s")' % table)]
        return [c for c in cols if c not in _META_COLS and c != "id"]

    def _refresh_counts(self):
        zh = self.conn.execute("SELECT COUNT(*) FROM zh_products").fetchone()[0]
        en = self.conn.execute("SELECT COUNT(*) FROM en_products").fetchone()[0]
        w = self.conn.execute("SELECT COUNT(*) FROM wide_rows").fetchone()[0]
        self.tree.item("zh", text="中文库 (%d 型号)" % zh)
        self.tree.item("en", text="英文库 (%d 型号)" % en)
        self.tree.item("wide", text="宽表库 (%d 条原始表述)" % w)

    def _refresh(self):
        if not self.conn:
            return
        self._loaded_models.clear()
        for iid in ("zh", "en", "wide"):
            for c in self.tree.get_children(iid):
                self.tree.delete(c)
        self._refresh_counts()
        self.detail.info_var.set("请选择左侧节点")
        self.detail.meta_var.set("")
        self.detail.sheet.clear()
        self._status("已刷新")

    def _open_db(self):
        path = filedialog.askopenfilename(
            title="打开正式筛选库", initialdir=str(DEFAULT_DB.parent),
            filetypes=[("SQLite", "*.db"), ("所有文件", "*.*")])
        if path:
            self.open_db(path)

    # ---------- 树加载 (懒加载) ----------

    def _on_open(self, ev):
        iid = self.tree.focus()
        if not iid or not self.conn:
            return
        if iid == "zh":
            self._load_models("zh", "zh_products")
        elif iid == "en":
            self._load_models("en", "en_products")
        elif iid == "wide":
            self._load_models("wide", None)
        elif iid.startswith("zh:") or iid.startswith("en:"):
            # 型号节点 (2段): 建节分组并加载节下字段; 节节点已含字段子节点, 无需再载
            if len(iid.split(":")) == 2:
                self._load_fields(iid)
        elif iid.startswith("wide:"):
            if len(iid.split(":")) == 2:
                self._load_wide_records(iid)

    def _load_models(self, prefix, table):
        """加载某库下的型号节点. 宽表按 distinct model."""
        if not self.conn:
            return
        if self.tree.get_children(prefix):
            return
        if table:
            first_col = self._std_cols(table)[0] if self._std_cols(table) else "model"
            rows = self.conn.execute(
                'SELECT model, "%s" FROM %s WHERE model != \'\' ORDER BY model' % (first_col, table)).fetchall()
            for r in rows:
                extra = "产品: %s" % str(r[1])[:20] if r[1] else ""
                self.tree.insert(prefix, "end", iid="%s:%s" % (prefix, r[0]),
                                 text=r[0], values=(extra,))
        else:
            rows = self.conn.execute(
                "SELECT DISTINCT model FROM wide_rows WHERE model != '' ORDER BY model").fetchall()
            for r in rows:
                self.tree.insert(prefix, "end", iid="wide:%s" % r[0], text=r[0], values=("",))

    def _load_fields(self, model_iid):
        """型号下按 16 节分组建节节点, 节内挂字段 (不拆散 section)."""
        if model_iid in self._loaded_models:
            return
        self._loaded_models.add(model_iid)
        if not self.conn:
            return
        prefix, model = model_iid.split(":", 1)
        table = "zh_products" if prefix == "zh" else "en_products"
        cols = self._zh_cols if prefix == "zh" else self._en_cols
        if not cols:
            return
        row = self.conn.execute("SELECT * FROM %s WHERE model=?" % table, (model,)).fetchone()
        if not row:
            return
        # 按节分组 (保留节内字段)
        by_sec: dict[int, list] = {}
        for c in cols:
            val = row[c]
            if not val or not str(val).strip():
                continue
            sec, _, label = c.partition("_")
            try:
                snum = int(sec)
            except ValueError:
                snum = 0
            by_sec.setdefault(snum, []).append((c, label, str(val)))
        for snum in sorted(by_sec):
            sec_iid = "%s:S%s" % (model_iid, snum)
            fields = by_sec[snum]
            self.tree.insert(model_iid, "end", iid=sec_iid,
                             text="第%d节 (%d 字段)" % (snum, len(fields)), open=False)
            for c, label, val in fields:
                self.tree.insert(sec_iid, "end", iid="%s:%s" % (sec_iid, c),
                                 text=label, values=(val[:30],))

    def _load_wide_records(self, model_iid):
        """宽表型号下按节分组, 节内挂原始表述记录."""
        if model_iid in self._loaded_models:
            return
        self._loaded_models.add(model_iid)
        if not self.conn:
            return
        model = model_iid.split(":", 1)[1]
        rows = self.conn.execute(
            "SELECT id, section, std_label, raw_label, vendor, seq, value FROM wide_rows "
            "WHERE model=? ORDER BY section, std_label, seq", (model,)).fetchall()
        by_sec: dict[int, list] = {}
        for r in rows:
            snum = r["section"] or 0
            by_sec.setdefault(snum, []).append(r)
        for snum in sorted(by_sec):
            sec_iid = "%s:S%s" % (model_iid, snum)
            recs = by_sec[snum]
            self.tree.insert(model_iid, "end", iid=sec_iid,
                             text="第%d节 (%d 记录)" % (snum, len(recs)), open=False)
            for r in recs:
                std = r["std_label"] or "(未识别)"
                txt = std
                self.tree.insert(sec_iid, "end", iid="%s:R%s" % (sec_iid, r["id"]), text=txt,
                                 values=(str(r["value"])[:28],))

    # ---------- 选择显示 ----------

    def _on_select(self, _ev=None):
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        if not self.conn:
            return
        parts = iid.split(":")
        # 一级: 表根
        if iid in ("zh", "en", "wide"):
            self._show_root(iid)
            return
        # 二级: 型号
        if len(parts) == 2:
            if parts[0] == "zh":
                self._show_unified_model("zh", parts[1], "zh_products")
            elif parts[0] == "en":
                self._show_unified_model("en", parts[1], "en_products")
            else:
                self._show_wide_model(parts[1])
            return
        # 三级: 节节点 (末段 S<n>)
        if len(parts) == 3 and parts[-1].startswith("S") and parts[-1][1:].isdigit():
            self._show_section_node(iid)
            return
        # 更深: 字段 / 宽表记录
        self._show_single(iid)

    def _show_root(self, iid):
        if iid == "zh":
            self._load_models("zh", "zh_products")
            n = self.conn.execute("SELECT COUNT(*) FROM zh_products").fetchone()[0]
            self.detail.show_model("中文库", "共 %d 个型号, 展开查看标准字段 (表格形式)" % n,
                                   {1: [("说明", "点左侧型号 → 本表按 16 节渲染该型号标准字段")]})
        elif iid == "en":
            self._load_models("en", "en_products")
            n = self.conn.execute("SELECT COUNT(*) FROM en_products").fetchone()[0]
            self.detail.show_model("英文库", "共 %d 个型号, 展开查看标准字段 (表格形式)" % n,
                                   {1: [("说明", "点左侧型号 → 本表按 16 节渲染该型号标准字段")]})
        else:
            self._load_models("wide", None)
            m = self.conn.execute("SELECT COUNT(DISTINCT model) FROM wide_rows").fetchone()[0]
            n = self.conn.execute("SELECT COUNT(*) FROM wide_rows").fetchone()[0]
            self.detail.show_wide("宽表库", "%d 个型号 / %d 条原始表述" % (m, n),
                                  {1: [("说明", "点左侧型号 → 本表按 16 节渲染全部原始表述 (标准标题|原始标题|值|品牌)", "", "")]})

    def _show_unified_model(self, prefix, model, table):
        row = self.conn.execute("SELECT * FROM %s WHERE model=?" % table, (model,)).fetchone()
        if not row:
            self.detail.info_var.set("未找到型号 %s" % model)
            return
        cols = self._zh_cols if table == "zh_products" else self._en_cols
        sections: dict[int, list] = {}
        n = 0
        for c in cols:
            val = row[c]
            if not val:
                continue
            sec, _, label = c.partition("_")
            try:
                snum = int(sec)
            except ValueError:
                snum = 0
            sections.setdefault(snum, []).append((label, str(val)))
            n += 1
        meta = "品牌来源: %s    有效字段: %d    %s" % (
            row["vendors"] or "-", n, row["created_at"] or "")
        self.detail.show_model(model, meta, sections)
        self._load_fields("%s:%s" % (prefix, model))

    def _show_wide_model(self, model):
        rows = self.conn.execute(
            "SELECT section, std_label, raw_label, value, vendor FROM wide_rows "
            "WHERE model=? ORDER BY section, std_label, seq", (model,)).fetchall()
        sections: dict[int, list] = {}
        for r in rows:
            snum = r["section"] or 0
            sections.setdefault(snum, []).append(
                (r["std_label"] or "-", r["raw_label"] or "-", r["value"] or "", r["vendor"] or "-"))
        meta = "共 %d 条原始表述 (宽表库, 保留原始)" % len(rows)
        self.detail.show_wide(model, meta, sections)
        self._load_wide_records("wide:%s" % model)

    def _show_section_node(self, iid):
        """三级节节点: 渲染该型号该节的全部字段 (统一库) / 记录 (宽表)."""
        parts = iid.split(":")
        if len(parts) != 3:
            return
        root, model, snode = parts
        snum = int(snode[1:])
        if root in ("zh", "en"):
            table = "zh_products" if root == "zh" else "en_products"
            cols = self._zh_cols if root == "zh" else self._en_cols
            row = self.conn.execute("SELECT * FROM %s WHERE model=?" % table, (model,)).fetchone()
            if not row:
                return
            rows = []
            for c in cols:
                val = row[c]
                if not val:
                    continue
                sec, _, label = c.partition("_")
                try:
                    sn = int(sec)
                except ValueError:
                    sn = 0
                if sn == snum:
                    rows.append((label, str(val)))
            self.detail.show_model(model, "第%d节 · %d 字段" % (snum, len(rows)),
                                   {snum: rows})
            # 展开该节下字段 (保证三级可选到字段)
            self._load_fields("%s:%s" % (root, model))
        else:  # 宽表节
            rows = self.conn.execute(
                "SELECT std_label, raw_label, value, vendor FROM wide_rows "
                "WHERE model=? AND section=? ORDER BY std_label, seq", (model, snum)).fetchall()
            data = [(r["std_label"] or "-", r["raw_label"] or "-",
                     r["value"] or "", r["vendor"] or "-") for r in rows]
            self.detail.show_wide(model, "第%d节 · %d 记录" % (snum, len(data)), {snum: data})
            self._load_wide_records("wide:%s" % model)

    def _show_single(self, iid):
        parts = iid.split(":")
        # 统一库字段: [zh|en, model, S3, col]
        if len(parts) == 4 and parts[0] in ("zh", "en"):
            prefix, model, snode, col = parts
            table = "zh_products" if prefix == "zh" else "en_products"
            row = self.conn.execute('SELECT "%s" FROM %s WHERE model=?' % (col, table), (model,)).fetchone()
            val = dict(row).get(col, "") if row else ""
            sec, _, label = col.partition("_")
            self.detail.show_model(model, "选中字段 [S%s] %s" % (sec, label),
                                   {1: [(label, str(val))]})
        # 宽表记录: [wide, model, S3, R<id>]
        elif len(parts) == 4 and parts[0] == "wide":
            model = parts[1]
            rid = parts[3][1:]
            r = self.conn.execute(
                "SELECT model, section, std_label, raw_label, value, vendor, seq FROM wide_rows WHERE id=?",
                (int(rid),)).fetchone()
            if r:
                self.detail.show_wide(r["model"], "选中记录 · 原始标题: %s · 品牌: %s · 序: %s" % (
                    r["raw_label"] or "-", r["vendor"] or "-", r["seq"] or "-"),
                    {r["section"] or 0: [(r["std_label"] or "-", r["raw_label"] or "-",
                                          r["value"] or "", r["vendor"] or "-")]})

    # ---------- 检索 ----------

    def _apply_search(self):
        if not self.conn:
            return
        q = self.search_var.get().strip().upper()
        if not q:
            self._refresh()
            return
        for root, table in (("zh", "zh_products"), ("en", "en_products"), ("wide", None)):
            for c in self.tree.get_children(root):
                self.tree.delete(c)
            if table:
                rows = self.conn.execute(
                    "SELECT model FROM %s WHERE upper(model) LIKE ? ORDER BY model" % table,
                    ("%%%s%%" % q,)).fetchall()
            else:
                rows = self.conn.execute(
                    "SELECT DISTINCT model FROM wide_rows WHERE upper(model) LIKE ? ORDER BY model",
                    ("%%%s%%" % q,)).fetchall()
            for r in rows:
                self.tree.insert(root, "end", iid="%s:%s" % (root, r[0]), text=r[0], values=("",))
        self._status("检索 '%s': 已过滤型号" % q)

    # ---------- 导出 ----------

    def _export_current(self):
        if not self.conn:
            messagebox.showwarning("未打开库", "请先打开正式筛选库")
            return
        path = filedialog.asksaveasfilename(
            title="导出当前详情", defaultextension=".csv", initialfile="筛选库视图.csv",
            filetypes=[("CSV", "*.csv")])
        if not path:
            return
        import csv
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["节", "标准标题", "原始标题", "值", "品牌"])
            for snum, rows in sorted(self.detail._sections.items()):
                for r in rows:
                    w.writerow([snum] + list(r))
        self._status("已导出: %s" % path)
        messagebox.showinfo("导出成功", "已保存:\n%s" % path)