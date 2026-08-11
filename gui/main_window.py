# -*- coding: utf-8 -*-
"""主窗口: 导入 → 读取 → 显示 → 覆写指向 四段式流程."""
from __future__ import annotations

import json
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from core.compare import compare
from core.docx_reader import TEMPLATE_PATH, read_msds
from core.structure import ParseResult

from .compare_view import CompareView
from .db_view import DbView  # 原总库视图保留 (不删, 供回退)
from .screened_view import DEFAULT_DB as _DEFAULT_SCREENED_DB
from .screened_view import ScreenedView
from .section_tree import SectionTree, SectionView
from .theme import COLOR_ACCENT, COLOR_BG, COLOR_BORDER, COLOR_PANEL, COLOR_TEXT


# 版本指纹: 每次大版本变更时递增, 显示在窗口标题与工具栏徽章, 便于确认运行的是最新版
_APP_VERSION = "v3.3"


class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"MSDS 结构读取 · 导入 - 读取 - 显示 - 覆写指向 / MSDS 总库 [{_APP_VERSION}]")
        self.geometry("1380x860")
        self.configure(bg=COLOR_BG)
        self.minsize(1100, 700)

        self.template: ParseResult | None = None
        self.product: ParseResult | None = None
        # 手动标注的字段权限: {(节, kind, index): 可编辑bool}; 未标注时用数据模型默认
        self._editable_overrides: dict[tuple[int, str, int], bool] = {}
        self._current_section = 1
        # 当前显示源: 'template' | 'product' (导入产品后显示产品; 恢复默认模板后切回模板)
        self._display_source = "template"

        self._build_toolbar()
        self._build_body()
        self._build_statusbar()

        # 启动时自动加载标准模板
        self._load_template()
        # 启动时自动打开正式筛选库 (若存在)
        if _DEFAULT_SCREENED_DB.exists():
            self.screened.open_db(str(_DEFAULT_SCREENED_DB))
            # 打开库后默认切到「正式筛选库」Tab (三表父子级目录)
            self.notebook.select(2)

    # ---------- UI 构建 ----------

    def _build_toolbar(self):
        bar = tk.Frame(self, bg=COLOR_PANEL, padx=8, pady=6, highlightthickness=1,
                       highlightbackground=COLOR_BORDER)
        bar.pack(fill="x")

        def btn(text, cmd, bg=COLOR_ACCENT, fg="white"):
            b = tk.Button(bar, text=text, command=cmd, bg=bg, fg=fg, relief="flat",
                          padx=14, pady=4, font=("Microsoft YaHei", 10), cursor="hand2")
            b.pack(side="left", padx=(0, 8))
            return b

        btn("📥 导入模板", self._pick_template)
        btn("📄 导入产品 MSDS", self._pick_product)
        btn("🔄 覆写指向分析", self._run_compare)
        tk.Button(bar, text="📤 导出 JSON", command=self._export_json,
                  bg="#E8EAED", fg=COLOR_TEXT, relief="flat", padx=14, pady=4,
                  font=("Microsoft YaHei", 10)).pack(side="left", padx=(0, 8))
        tk.Button(bar, text="↩️ 恢复默认模板", command=self._restore_default_template,
                  bg="#E8EAED", fg=COLOR_TEXT, relief="flat", padx=14, pady=4,
                  font=("Microsoft YaHei", 10)).pack(side="left", padx=(0, 8))

        self.tpl_var = tk.StringVar(value="模板: 未加载")
        tk.Label(bar, textvariable=self.tpl_var, bg=COLOR_PANEL, fg=COLOR_TEXT,
                 font=("Microsoft YaHei", 9)).pack(side="right", padx=(8, 0))
        self.prod_var = tk.StringVar(value="产品: 未导入")
        tk.Label(bar, textvariable=self.prod_var, bg=COLOR_PANEL, fg=COLOR_TEXT,
                 font=("Microsoft YaHei", 9)).pack(side="right")
        # 版本徽章: 确认运行版本 (旧代码无此徽章)
        tk.Label(bar, text=f"✔ {_APP_VERSION}", bg="#1E8E3E", fg="white",
                 font=("Microsoft YaHei", 9, "bold"), padx=8, pady=2,
                 cursor="hand2").pack(side="right", padx=(12, 4))

    def _build_body(self):
        body = tk.Frame(self, bg=COLOR_BG)
        body.pack(fill="both", expand=True)

        # 左侧导航
        self.nav = SectionTree(body, on_select=self._show_section)
        self.nav.pack(side="left", fill="y", padx=(6, 3), pady=6)

        # 右侧 Notebook: 结构显示 / 覆写指向
        self.notebook = ttk.Notebook(body)
        self.notebook.pack(side="left", fill="both", expand=True, padx=(3, 6), pady=6)

        # Tab1: 16节结构化显示
        self.section_view = SectionView(self.notebook)
        self.notebook.add(self.section_view, text="16 节结构")

        # Tab2: 覆写指向
        self.compare_view = CompareView(self.notebook)
        self.notebook.add(self.compare_view, text="覆写指向分析")

        # Tab3: 正式筛选库 (中文/英文/宽表 三表父子级目录)
        self.screened = ScreenedView(self.notebook, on_status=self._db_status)
        self.notebook.add(self.screened, text="正式筛选库")

        # 切到总库 Tab 时隐藏主窗口左侧 16 节导航树 (双目录原则: 总库内 左=型号 右=节)
        self.notebook.bind("<<NotebookTabChanged>>", self._on_main_tab_changed)

    def _build_statusbar(self):
        self.status_var = tk.StringVar(value="就绪")
        bar = tk.Frame(self, bg=COLOR_PANEL, highlightthickness=1, highlightbackground=COLOR_BORDER)
        bar.pack(fill="x", side="bottom")
        tk.Label(bar, textvariable=self.status_var, bg=COLOR_PANEL, fg="#5F6368",
                 font=("Microsoft YaHei", 9), anchor="w", padx=10, pady=3).pack(fill="x")

    def _db_status(self, msg: str):
        """总库 Tab 状态 → 主窗口状态栏."""
        self.status_var.set(msg)

    def _on_main_tab_changed(self, _ev):
        """双目录: 切到「正式筛选库」Tab 时隐藏主窗口左侧 16 节导航树,
        避免 导航树|型号|字段 三目录重叠; 其他 Tab 恢复显示."""
        try:
            idx = self.notebook.index(self.notebook.select())
        except (tk.TclError, ValueError):
            return
        if idx == 2:   # 正式筛选库 (三表父子级目录自带左树)
            self.nav.pack_forget()
        else:
            self.nav.pack(side="left", fill="y", padx=(6, 3), pady=6)

    # ---------- 动作 ----------

    def _load_template(self, path: Path | None = None):
        path = path or TEMPLATE_PATH
        if not path.exists():
            self.status_var.set(f"⚠️ 标准模板不存在: {path}")
            return
        try:
            self.template = read_msds(path)
            self.tpl_var.set(f"模板: {path.name}")
            self._display_source = "template"   # 导入/重载模板 → 显示模板
            self.nav.set_result(self.template)
            self.nav.select(1)
            self._show_section(1)
            s = self.template.summary()
            self.status_var.set(
                f"✅ 模板已加载: {s['sections']}节 / {s['tables']}表 / {s['fields']}字段 "
                f"/ {s['components']}成分 / {s['anomalies']}异常"
            )
        except Exception as exc:
            messagebox.showerror("模板加载失败", str(exc))

    def _pick_template(self):
        path = filedialog.askopenfilename(title="选择 MSDS 模板", filetypes=[("Word 文档", "*.docx *.doc"), ("所有文件", "*.*")])
        if path:
            self._load_template(Path(path))

    def _pick_product(self):
        path = filedialog.askopenfilename(title="选择产品 MSDS", filetypes=[("Word 文档", "*.docx *.doc"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            self.product = read_msds(path)
            self.prod_var.set(f"产品: {Path(path).name}")
            self._editable_overrides.clear()   # 新产品导入后, 字段权限复位为默认
            self._display_source = "product"   # 导入产品 → 显示产品
            # 读取后立即显示产品结构
            self.nav.set_result(self.product)
            self.nav.select(1)
            self._show_section(1)
            s = self.product.summary()
            self.status_var.set(
                f"✅ 产品已读取: {s['sections']}节 / {s['fields']}字段 / {s['components']}成分 "
                f"/ {s['anomalies']}异常  (点击『覆写指向分析』标注可覆写字段)"
            )
            if s["anomalies"]:
                self._show_anomalies(self.product)
        except Exception as exc:
            messagebox.showerror("读取失败", str(exc))

    def _show_section(self, num: int):
        self._current_section = num
        src = self._display_source_of()
        if src:
            self.section_view.show_section(num, src, self._editable_overrides, self._toggle_editable)

    def _display_source_of(self) -> ParseResult | None:
        """当前显示源: 按显示状态返回 产品/模板 (恢复默认模板后显示模板)."""
        if self._display_source == "product" and self.product:
            return self.product
        return self.template

    def _restore_default_template(self):
        """恢复显示默认模板 (内化副本 MSDS_CN 国彩 模板.docx), 并重置比对基准.

        真实功能: 重新读取内化模板文件覆盖 self.template, 显示源切回模板,
        字段权限标注复位为模板默认. 已导入的产品保留 (比对仍可用), 但界面
        回到 16 节结构页显示模板内容.
        """
        if not TEMPLATE_PATH.exists():
            self.status_var.set(f"⚠️ 默认模板不存在: {TEMPLATE_PATH}")
            return
        try:
            self.template = read_msds(TEMPLATE_PATH)
            self.tpl_var.set(f"模板: {TEMPLATE_PATH.name} (默认)")
            self._display_source = "template"
            self._editable_overrides.clear()
            self.nav.set_result(self.template)
            self.nav.select(1)
            self._show_section(1)
            self.notebook.select(0)   # 切回 16 节结构页
            s = self.template.summary()
            self.status_var.set(
                f"✅ 已恢复显示默认模板: {s['sections']}节 / {s['tables']}表 / "
                f"{s['fields']}字段 / {s['components']}成分 / {s['anomalies']}异常"
            )
        except Exception as exc:
            messagebox.showerror("恢复默认模板失败", str(exc))

    def _toggle_editable(self, key: tuple[int, str, int], new_state: bool):
        """用户点击徽章: 记录该字段的手动标注."""
        self._editable_overrides[key] = new_state
        sec, kind, idx = key
        self.status_var.set(
            f"已标注字段权限: 第{sec}节 [{kind}#{idx}] → {'可编辑' if new_state else '不可编辑'}"
        )

    def _run_compare(self):
        if not self.template:
            messagebox.showwarning("缺少模板", "请先导入模板")
            return
        if not self.product:
            messagebox.showwarning("缺少产品", "请先导入产品 MSDS")
            return
        cr = compare(self.template, self.product)
        # 覆写建议: 模板残留风险(review)字段默认锁为不可编辑, 防误覆盖; 用户可手动改
        self._apply_review_locks(cr)
        self.compare_view.show_compare(cr)
        self._show_section(self._current_section)
        self.notebook.select(1)   # 切到比对页展示自动建议
        self.status_var.set(
            f"✅ 覆写指向分析完成: 共 {len(cr.decisions)} 字段, {len(cr.residue_risks)} 处残留风险 — "
            f"比对建议见本页, 字段权限请到『16 节结构』页点击徽章手动标注"
        )

    def _apply_review_locks(self, cr):
        """把 review (模板有旧值但产品未提供) 的字段默认设为不可编辑."""
        for d in cr.residue_risks:
            sec = self.template.sections.get(d.section)
            if not sec:
                continue
            for i, f in enumerate(sec.fields):
                if f.label == d.label:
                    key = (d.section, "field", i)
                    if key not in self._editable_overrides:
                        self._editable_overrides[key] = False
                    break

    def _show_anomalies(self, result: ParseResult):
        msg = "\n".join(f"[{'⚠️' if a.level=='warn' else '❌'}] S{a.section}: {a.message}"
                        for a in result.anomalies[:12])
        if len(result.anomalies) > 12:
            msg += f"\n... 等{len(result.anomalies)}项"
        if result.anomalies:
            messagebox.showwarning(f"读取告警 ({len(result.anomalies)}项)", msg)

    def _export_json(self):
        src = self._display_source_of()
        if not src:
            messagebox.showwarning("无内容", "请先导入文件")
            return
        default = Path("outputs") / f"{Path(src.file_name).stem}_读取结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        path = filedialog.asksaveasfilename(
            title="导出 JSON", defaultextension=".json", initialfile=default.name,
            filetypes=[("JSON", "*.json")], initialdir=str(default.parent))
        if not path:
            return
        data = {
            "file": src.file_name,
            "sha256": src.sha256,
            "header": src.header,
            "footer": src.footer,
            "sections": {},
        }
        for num, sec in src.sections.items():
            # 统一行模型 (节标题 | 字段 | 子标题 | 说明), 每行带 editable 字段权限
            rows = []
            for row in sec.iter_rows():
                if row.kind == "section":
                    continue
                rows.append({
                    "kind": row.kind,
                    "seq": row.seq,
                    "label": row.label,
                    "value": row.value,
                    "editable": self._editable_overrides.get(
                        (num, row.kind, row.index), row.editable),
                })
            data["sections"][str(num)] = {
                "title": sec.full_title,
                "rows": rows,   # 标题/内容两列 + 可编辑状态 (供覆写/编辑流程)
                "fields": {f.label: f.value for f in sec.fields},
                "lines": sec.lines,
                "components": [
                    {"name": c.name, "cas": c.cas, "conc": c.conc,
                     "editable": self._editable_overrides.get(
                         (num, "component", i), c.editable)}
                    for i, c in enumerate(sec.components)
                ],
            }
        if self.template and self.product:
            cr = compare(self.template, self.product)
            data["overwrite_analysis"] = {
                "template": self.template.file_name,
                "product": self.product.file_name,
                "decisions": [
                    {"section": d.section, "label": d.label,
                     "template_value": d.template_value, "product_value": d.product_value,
                     "write_source": d.write_source, "reason": d.reason}
                    for d in cr.decisions
                ],
            }
        Path(path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        self.status_var.set(f"✅ 已导出: {path}")
        messagebox.showinfo("导出成功", f"已保存到:\n{path}")
