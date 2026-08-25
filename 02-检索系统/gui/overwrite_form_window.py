# -*- coding: utf-8 -*-
"""固定冠志模板的 17 节 GUI 表单、覆写与 Word 产出。"""
from __future__ import annotations

import json
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk
import tkinter as tk
from tkinter.scrolledtext import ScrolledText

from core.docx_reader import TEMPLATE_PATH, read_msds
from core.schema import SECTION_SCHEMAS, standard_name
from core.edit_policy import (
    is_editable_row,
    normalize_edit_label,
    validate_write_items_permissions,
)

BASE_DIR = Path(__file__).resolve().parents[2]
PICTOGRAM_DIR = BASE_DIR / "04-推断引擎" / "推断引擎程序"
if str(PICTOGRAM_DIR) not in sys.path:
    sys.path.insert(0, str(PICTOGRAM_DIR))
from pictogram_registry import (  # noqa: E402
    asset_bytes, display_value, extract_codes, label as pictogram_label,
    ordered_codes,
)

ENGINE_DIR = BASE_DIR / "05-覆写模块"
if str(ENGINE_DIR) not in sys.path:
    sys.path.insert(0, str(ENGINE_DIR))

from msds_form_overwrite import (  # noqa: E402
    EditableField, FormState, build_write_items,
)
import msds_overwrite_engine as overwrite_engine  # noqa: E402


@dataclass
class UiField:
    section: int
    seq: str
    label: str
    parent: str
    current: str
    editable: bool
    kind: str
    key: str | None = None
    semantic: str = ""


S9_SCHEMA_FIELDS = [field for field in SECTION_SCHEMAS[9] if field.kind == "field"]
S9_SCHEMA_BY_NAME = {field.name: field for field in S9_SCHEMA_FIELDS}


class OverwriteFormWindow(tk.Toplevel):
    """在固定模板上编辑白名单字段并产出 Word。"""

    def __init__(self, master, template_path: Path = TEMPLATE_PATH):
        super().__init__(master)
        self.title("MSDS 17 节固定模板覆写与产出")
        self.geometry("1280x860")
        self.minsize(1080, 700)
        self.configure(bg="#F5F7FA")
        self.template_path = Path(template_path).resolve()
        if self.template_path != Path(TEMPLATE_PATH).resolve():
            raise ValueError("覆写表单只能使用项目批准的固定冠志中文模板")
        self.result = read_msds(self.template_path)
        self.state = FormState(template=self.template_path, result=self.result)
        self.fields_by_section: dict[int, list[UiField]] = {i: [] for i in range(17)}
        self._widgets: dict[str, ScrolledText] = {}
        self._component_vars: list[dict[str, tk.StringVar]] = []
        self._bio_vars: list[list[tk.StringVar]] = []
        self._s15_controls: list[tuple[str, tk.StringVar, str]] = []
        self._s9_alias_vars: dict[str, tk.StringVar] = {}
        self._pictogram_vars: dict[str, dict[str, tk.BooleanVar]] = {}
        self._pictogram_touched: set[str] = set()
        self._pictogram_photos: list = []
        self._s9_alias_values: dict[str, list[str]] = {
            field.name: [field.name, *field.aliases] for field in S9_SCHEMA_FIELDS
        }
        self._scroll_bindings: list[tuple[str, str]] = []
        self._current_section = 1
        self._build_catalog()
        self._build_ui()
        self._show_section(1)
        self._scroll_bindings = [
            ("<MouseWheel>", self.bind_all("<MouseWheel>", self._on_mousewheel, add="+")),
            ("<Button-4>", self.bind_all("<Button-4>", lambda _e: self._scroll_canvas(-3), add="+")),
            ("<Button-5>", self.bind_all("<Button-5>", lambda _e: self._scroll_canvas(3), add="+")),
        ]

    # ---------- 权限与目录 ----------
    def _note_semantic(self, section: int, value: str, parent: str) -> tuple[str, str]:
        text = str(value or "").strip()
        if section == 11:
            if text.startswith("该产品无可用"):
                return "产品说明", "normal"
            if text.startswith("以下为"):
                return "成分参考引导段", "normal"
        if section == 12:
            if text.startswith("该产品无可用"):
                return "产品说明", "normal"
            if text.startswith("以下为"):
                return "成分参考引导段", "normal"
        if section == 15:
            if "物质或混合物" in text or "安全、健康和环保" in text:
                return "指引段", "guide"
            return "法规条目", "law"
        if section == 16:
            # S16 的 note 正文必须绑定数据库标准字段“免责声明”，而不是
            # 把整段旧模板正文当作字段标签，否则 GUI 无法稳定回传修改值。
            return "免责声明", "normal"
        # S13 没有普通标签，使用原通栏文本作为定位标签，覆写引擎可精确命中。
        return text, "normal"

    def _build_catalog(self) -> None:
        """从固定模板原始行建立完整可见目录，权限只走白名单。"""
        for sec_no in range(17):
            sec = self.result.sections.get(sec_no)
            if not sec:
                continue
            parent = ""
            occurrences: dict[tuple[str, str], int] = {}
            for row in sec.iter_rows():
                if row.kind == "sub":
                    parent = f"{row.seq} {row.label}".strip()
                    self.fields_by_section[sec_no].append(UiField(
                        sec_no, row.seq, row.label, parent, "", False, "sub"))
                    continue
                if row.kind == "subtable":
                    self.fields_by_section[sec_no].append(UiField(
                        sec_no, row.seq, row.label, parent, "", False, "subtable"))
                    continue
                if row.kind not in {"field", "note"}:
                    continue
                raw_label = str(row.label or "").strip()
                label = normalize_edit_label(raw_label) or "通栏说明"
                display_seq = row.seq
                if sec_no == 9 and label:
                    # S9 以飞书 Schema 的 37 个标准字段为目录；源文件中的
                    # pH值（1%水溶液）等特殊写法只做别名，不改变标准标签。
                    label = standard_name(9, label) or label
                    schema_index = next((i for i, item in enumerate(S9_SCHEMA_FIELDS, 1)
                                         if item.name == label), None)
                    if schema_index:
                        display_seq = f"9.{schema_index}"
                semantic = ""
                if row.kind == "note":
                    if sec_no == 15 and label in {"其它的规定", "符合下列法规要求"}:
                        # 两个单格 note 是 S15 的正文输入槽位；结构首行仍由
                        # 覆写引擎保留，正文通过专用控件写入下一行。
                        value = str(row.value or "").strip()
                        self.state.s15_special[label] = "" if value == label else value
                        continue
                    label, semantic = self._note_semantic(sec_no, row.value, parent)
                    if sec_no == 15 and semantic == "law":
                        # S15 法规正文统一在专用动态列表中展示，避免同时进入普通
                        # write_items 与法规列表而产生重复条目。
                        if str(row.value or "").strip():
                            self.state.laws.append(str(row.value).strip())
                        continue
                editable = is_editable_row(sec_no, row)
                # 序号/标签是模板结构；只有白名单字段允许编辑。
                pair = (display_seq or "", label)
                index = occurrences.get(pair, 0)
                occurrences[pair] = index + 1
                key = f"S{sec_no}|{display_seq}|{label}" + (f"#{index + 1}" if index else "")
                item = UiField(sec_no, display_seq, label, parent, str(row.value or ""),
                               editable, row.kind, key if editable else None, semantic)
                self.fields_by_section[sec_no].append(item)
                if editable:
                    self.state.fields[key] = EditableField(
                        section=sec_no, seq=display_seq, label=label,
                        parent=parent, current=str(row.value or ""), kind=row.kind, key=key)
                    self.state.values[key] = str(row.value or "")

            # 子表数据行权限：表头固定，数据行可在专用控件中编辑。
            if sec_no == 3:
                self.state.components = [
                    {"name": str(c.name or ""), "cas": str(c.cas or ""), "conc": str(c.conc or "")}
                    for c in sec.components
                ]
                product = next((x for x in self.fields_by_section[3] if x.label == "产品类型"), None)
                if product:
                    self.state.product_type = product.current or "混合物"
            if sec_no == 8:
                for row in sec.iter_rows():
                    if row.kind == "subtable" and row.label == "生物限值":
                        self.state.bio_rows = [list(x) for x in row.sub_rows]

            if sec_no == 9:
                existing = {field.label for field in self.fields_by_section[9] if field.kind in {"field", "note"}}
                # 补齐飞书规范定义但当前模板没有物理行的 S9 字段，供表单填写后
                # 由覆写引擎按模板参考行新增；9.37 其他信息始终置于末尾。
                missing = [field for field in S9_SCHEMA_FIELDS if field.name not in existing]
                if missing:
                    insert_at = next((i for i, item in enumerate(self.fields_by_section[9])
                                      if item.label == "其他信息"), len(self.fields_by_section[9]))
                    additions = []
                    for schema_index, schema_field in enumerate(S9_SCHEMA_FIELDS, 1):
                        if schema_field.name not in existing:
                            key = f"S9|9.{schema_index}|{schema_field.name}"
                            additions.append(UiField(9, f"9.{schema_index}", schema_field.name, "",
                                                     "", True, "field", key))
                            self.state.fields[key] = EditableField(
                                section=9, seq=f"9.{schema_index}", label=schema_field.name,
                                parent="", current="", kind="field", key=key)
                            self.state.values[key] = ""
                    self.fields_by_section[9][insert_at:insert_at] = additions

    # ---------- UI ----------
    def _build_ui(self) -> None:
        top = tk.Frame(self, bg="#FFFFFF", padx=10, pady=8)
        top.pack(fill="x")
        tk.Label(top, text="固定模板：PEA-4139 MSDS_CN 冠志 模板.docx",
                 bg="#FFFFFF", fg="#17365D", font=("Microsoft YaHei", 11, "bold")).pack(side="left")
        tk.Label(top, text="仅白名单字段可编辑，其余均为锁定状态",
                 bg="#FFFFFF", fg="#C62828", font=("Microsoft YaHei", 10)).pack(side="right")

        body = tk.Frame(self, bg="#F5F7FA")
        body.pack(fill="both", expand=True, padx=8, pady=8)
        nav_frame = tk.Frame(body, bg="#17365D", width=210)
        nav_frame.pack(side="left", fill="y")
        nav_frame.pack_propagate(False)
        tk.Label(nav_frame, text="17 节表单", bg="#17365D", fg="white",
                 font=("Microsoft YaHei", 12, "bold"), pady=10).pack(fill="x")
        self.nav_list = tk.Listbox(nav_frame, bg="#17365D", fg="white", selectbackground="#3976B8",
                                   relief="flat", borderwidth=0, font=("Microsoft YaHei", 10),
                                   activestyle="none")
        self.nav_list.pack(fill="both", expand=True, padx=6, pady=6)
        for i in range(17):
            self.nav_list.insert("end", f"S{i}  {self._section_title(i)}")
        self.nav_list.bind("<<ListboxSelect>>", self._on_nav)

        right = tk.Frame(body, bg="#F5F7FA")
        right.pack(side="left", fill="both", expand=True, padx=(8, 0))
        self.canvas = tk.Canvas(right, bg="#F5F7FA", highlightthickness=0)
        scrollbar = ttk.Scrollbar(right, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.inner = tk.Frame(self.canvas, bg="#F5F7FA")
        self.canvas_window = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", lambda _e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self.canvas_window, width=e.width))

        bottom = tk.Frame(self, bg="#FFFFFF", padx=10, pady=8)
        bottom.pack(fill="x")
        tk.Button(bottom, text="保存写入项 JSON", command=self._save_json,
                  bg="#607D8B", fg="white", relief="flat", padx=12, pady=6,
                  font=("Microsoft YaHei", 10)).pack(side="left", padx=(0, 8))
        tk.Button(bottom, text="确认覆写并产出 Word", command=self._overwrite,
                  bg="#0B8043", fg="white", relief="flat", padx=16, pady=6,
                  font=("Microsoft YaHei", 10, "bold")).pack(side="left")
        tk.Button(bottom, text="关闭", command=self._close_window,
                  bg="#E8EAED", fg="#333333", relief="flat", padx=12, pady=6,
                  font=("Microsoft YaHei", 10)).pack(side="right")
        self.status_var = tk.StringVar(value="请在左侧选择节；灰色字段不可编辑。")
        tk.Label(bottom, textvariable=self.status_var, bg="#FFFFFF", fg="#5F6368",
                 font=("Microsoft YaHei", 9)).pack(side="right", padx=12)

    def _section_title(self, num: int) -> str:
        from core.msds_db import SEC_TITLES
        return SEC_TITLES.get(num, f"第{num}节")

    def _on_nav(self, _event=None) -> None:
        selection = self.nav_list.curselection()
        if selection:
            self._show_section(int(selection[0]))

    def _clear_inner(self) -> None:
        for widget in self.inner.winfo_children():
            widget.destroy()
        self._widgets.clear()
        self._component_vars.clear()
        self._bio_vars.clear()
        self._s15_controls.clear()
        self._s9_alias_vars.clear()
        self._pictogram_vars.clear()
        self._pictogram_touched.clear()
        self._pictogram_photos.clear()

    def _scroll_canvas(self, units: int) -> None:
        self.canvas.yview_scroll(units, "units")

    def _on_mousewheel(self, event) -> None:
        delta = int(getattr(event, "delta", 0))
        if delta:
            self._scroll_canvas(-max(1, abs(delta) // 120) * (1 if delta > 0 else -1))

    def _add_s9_alias(self, field: UiField) -> None:
        """为当前 S9 标准字段增加本次表单会话可用的源标签别名。"""
        value = simpledialog.askstring("添加 S9 特殊表述", f"为“{field.label}”添加特殊表述：", parent=self)
        alias = str(value or "").strip()
        if not alias:
            return
        aliases = self._s9_alias_values.setdefault(field.label, [field.label])
        if alias not in aliases:
            aliases.append(alias)
        if field.key in self._s9_alias_vars:
            self._s9_alias_vars[field.key].set(alias)

    def _close_window(self) -> None:
        for event, binding in self._scroll_bindings:
            try:
                self.unbind_all(event, binding)
            except tk.TclError:
                pass
        self.destroy()

    def _make_text(self, parent, text: str, editable: bool, key: str | None, height=4):
        box = ScrolledText(parent, height=height, wrap="word", font=("Microsoft YaHei", 10),
                           relief="solid", borderwidth=1)
        box.insert("1.0", text)
        if not editable:
            box.configure(state="disabled", bg="#ECEFF1", fg="#6B7280")
        else:
            box.configure(bg="#FFFFFF", fg="#222222")
            if key:
                self._widgets[key] = box
        return box

    def _show_section(self, section: int) -> None:
        self._collect_current()
        self._current_section = section
        self._clear_inner()
        tk.Label(self.inner, text=f"S{section}  {self._section_title(section)}",
                 bg="#17365D", fg="white", anchor="w", padx=12, pady=9,
                 font=("Microsoft YaHei", 13, "bold")).pack(fill="x", pady=(0, 8))
        for field in self.fields_by_section[section]:
            if field.kind == "subtable":
                if section == 8:
                    self._render_bio_table()
                else:
                    self._render_locked_note("子表表头", "固定模板表头不可编辑")
                continue
            if field.kind == "sub":
                self._render_locked_note(f"{field.seq} {field.label}", "父级结构行不可编辑")
                continue
            self._render_field(field)
        if section == 3:
            self._render_components()
        if section == 8 and not any(x.kind == "subtable" for x in self.fields_by_section[8]):
            self._render_bio_table()
        if section == 15:
            self._render_s15_laws()
        self.canvas.yview_moveto(0)
        self.nav_list.selection_clear(0, "end")
        self.nav_list.selection_set(section)

    def _render_locked_note(self, label: str, value: str) -> None:
        frame = tk.Frame(self.inner, bg="#ECEFF1", padx=8, pady=5)
        frame.pack(fill="x", pady=3)
        tk.Label(frame, text=f"🔒 {label}", bg="#ECEFF1", fg="#606770",
                 font=("Microsoft YaHei", 10, "bold"), anchor="w").pack(fill="x")
        tk.Label(frame, text=value, bg="#ECEFF1", fg="#7A7F85",
                 font=("Microsoft YaHei", 9), anchor="w").pack(fill="x")

    def _render_field(self, field: UiField) -> None:
        frame = tk.Frame(self.inner, bg="#FFFFFF", padx=8, pady=6,
                         highlightthickness=1, highlightbackground="#DADCE0")
        frame.pack(fill="x", pady=3)
        title = f"{field.seq} {field.label}" if field.seq else field.label
        fg = "#17365D" if field.editable else "#6B7280"
        suffix = "  [可编辑]" if field.editable else "  [不可编辑]"
        tk.Label(frame, text=title + suffix, bg="#FFFFFF", fg=fg,
                 font=("Microsoft YaHei", 10, "bold"), anchor="w").pack(fill="x")
        if field.section == 2 and normalize_edit_label(field.label) in {"象形图", "GHS象形图"}:
            self._render_pictogram_field(frame, field)
            return
        if field.section == 9 and field.editable:
            alias_bar = tk.Frame(frame, bg="#FFFFFF")
            alias_bar.pack(fill="x", pady=(4, 0))
            tk.Label(alias_bar, text="特殊表述映射（不改模板标签）：",
                     bg="#FFFFFF", fg="#6B7280", font=("Microsoft YaHei", 8)).pack(side="left")
            alias_var = tk.StringVar(value=self.state.label_overrides.get(field.key, field.label))
            self._s9_alias_vars[field.key] = alias_var
            aliases = self._s9_alias_values.setdefault(field.label, [field.label])
            ttk.Combobox(alias_bar, textvariable=alias_var, values=aliases,
                         state="readonly", width=32, font=("Microsoft YaHei", 9)).pack(side="left", padx=4)
            tk.Button(alias_bar, text="添加表述", command=lambda f=field: self._add_s9_alias(f),
                      bg="#E8EAED", fg="#425466", relief="flat", padx=6,
                      font=("Microsoft YaHei", 8)).pack(side="left")
        # Section 切换会销毁并重建控件；可编辑字段必须从 FormState 读取最新值，
        # 不能继续使用建目录时保存的模板初始值，否则切回该节时会把用户输入覆盖掉。
        current = self.state.values.get(field.key, field.current) if field.editable else field.current
        box = self._make_text(frame, current, field.editable, field.key,
                              height=5 if len(current) > 90 else 3)
        box.pack(fill="x", pady=(4, 0))
        if field.kind == "note" and field.semantic:
            tk.Label(frame, text=f"定位语义：{field.semantic}", bg="#FFFFFF", fg="#8A8F98",
                     font=("Microsoft YaHei", 8), anchor="w").pack(fill="x")

    def _pictogram_photo(self, blob: bytes, ext: str):
        """把唯一标准 GHS 红菱形图片缩放为表单预览图。"""
        import io
        try:
            from PIL import Image, ImageTk
            image = Image.open(io.BytesIO(blob)).convert("RGBA")
            image.thumbnail((86, 86), Image.LANCZOS)
            photo = ImageTk.PhotoImage(image)
        except Exception:
            try:
                photo = tk.PhotoImage(data=blob)
            except Exception:
                return None
        self._pictogram_photos.append(photo)
        return photo

    def _render_pictogram_field(self, frame, field: UiField) -> None:
        """S2 象形图专用控件：代码选择、标准名称和红菱形预览三者绑定。"""
        current = self.state.values.get(field.key, field.current) if field.editable else field.current
        current_codes = set(extract_codes(current))
        controls: dict[str, tk.BooleanVar] = {}
        self._pictogram_vars[field.key] = controls
        palette = tk.Frame(frame, bg="#FFFFFF")
        palette.pack(fill="x", pady=(6, 2))
        for index in range(1, 10):
            code = f"GHS{index:02d}"
            var = tk.BooleanVar(value=code in current_codes)
            controls[code] = var
            blob_info = asset_bytes(code)
            photo = self._pictogram_photo(*blob_info) if blob_info else None
            item = tk.Frame(palette, bg="#FFFFFF", padx=4, pady=4,
                            highlightthickness=1, highlightbackground="#DADCE0")
            item.grid(row=(index - 1) // 3, column=(index - 1) % 3,
                      sticky="w", padx=3, pady=3)
            check = tk.Checkbutton(
                item, text=f"{code}  {pictogram_label(code)}", variable=var,
                image=photo, compound="top", anchor="center", justify="center",
                bg="#FFFFFF", activebackground="#EEF5FF", relief="flat",
                font=("Microsoft YaHei", 9),
                command=lambda key=field.key: self._pictogram_touched.add(key),
            )
            if photo is not None:
                check.image = photo
            check.pack()
        if current and not current_codes and current not in {"无", "无数据", "不适用"}:
            tk.Label(frame, text=f"未识别的原始表述：{current}（选择代码后将改为标准结果）",
                     bg="#FFF8E1", fg="#8A5A00", anchor="w",
                     font=("Microsoft YaHei", 9)).pack(fill="x", pady=(4, 0))
        tk.Label(frame, text="输出只写入 GHS 代码和标准名称；图片来自统一 GHS 红菱形图库。",
                 bg="#FFFFFF", fg="#6B7280", anchor="w",
                 font=("Microsoft YaHei", 8)).pack(fill="x", pady=(4, 0))

    def _render_components(self) -> None:
        frame = tk.LabelFrame(self.inner, text="S3 成分子表（数据行可编辑，表头固定）",
                              bg="#FFFFFF", fg="#17365D", font=("Microsoft YaHei", 10, "bold"),
                              padx=8, pady=8)
        frame.pack(fill="x", pady=6)
        labels = ("化学品名称", "CAS 编号", "含量 %（w/w）")
        for col, label in enumerate(labels):
            tk.Label(frame, text=label, bg="#E8EEF5", fg="#17365D",
                     font=("Microsoft YaHei", 9, "bold"), padx=4, pady=4).grid(row=0, column=col, sticky="ew")
            frame.grid_columnconfigure(col, weight=1)
        rows = list(self.state.components) + [{"name": "", "cas": "", "conc": ""} for _ in range(3)]
        for ri, row in enumerate(rows, start=1):
            vars_row = {k: tk.StringVar(value=row[k]) for k in ("name", "cas", "conc")}
            self._component_vars.append(vars_row)
            for ci, key in enumerate(("name", "cas", "conc")):
                entry = tk.Entry(frame, textvariable=vars_row[key], font=("Microsoft YaHei", 10))
                entry.grid(row=ri, column=ci, sticky="ew", padx=2, pady=2)

    def _render_bio_table(self) -> None:
        frame = tk.LabelFrame(self.inner, text="S8.2 生物限值（数据行可编辑，五列表头固定）",
                              bg="#FFFFFF", fg="#17365D", font=("Microsoft YaHei", 10, "bold"),
                              padx=8, pady=8)
        frame.pack(fill="x", pady=6)
        for col, label in enumerate(("组分名称", "标准来源", "生物监测指标", "生物限值", "采样时间")):
            tk.Label(frame, text=label, bg="#E8EEF5", fg="#17365D",
                     font=("Microsoft YaHei", 8, "bold"), padx=3, pady=4).grid(row=0, column=col, sticky="ew")
            frame.grid_columnconfigure(col, weight=1)
        rows = list(self.state.bio_rows) + [[""] * 5 for _ in range(2)]
        for ri, row in enumerate(rows, start=1):
            vars_row = [tk.StringVar(value=(row[i] if i < len(row) else "")) for i in range(5)]
            self._bio_vars.append(vars_row)
            for ci, var in enumerate(vars_row):
                tk.Entry(frame, textvariable=var, font=("Microsoft YaHei", 9)).grid(
                    row=ri, column=ci, sticky="ew", padx=2, pady=2)

    def _render_s15_laws(self) -> None:
        frame = tk.LabelFrame(self.inner, text="S15 法规信息（说明槽位与法规条目可编辑；节结构固定）",
                              bg="#FFFFFF", fg="#17365D", font=("Microsoft YaHei", 10, "bold"),
                              padx=8, pady=8)
        frame.pack(fill="x", pady=6)
        # 结构首行锁定；正文说明与法规条目使用专用输入控件，避免普通
        # field/note 白名单把 S15 输入误判为不存在。
        tk.Label(frame, text="🔒 指引段及节标题固定；以下两个说明槽位写入标题下方正文",
                 bg="#ECEFF1", fg="#6B7280", font=("Microsoft YaHei", 9),
                 anchor="w").grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 5))
        row_index = 1
        for label in ("其它的规定", "符合下列法规要求"):
            var = tk.StringVar(value=self.state.s15_special.get(label, ""))
            self._s15_controls.append((label, var, "special"))
            tk.Label(frame, text=label, bg="#FFFFFF", fg="#425466",
                     font=("Microsoft YaHei", 9), anchor="w").grid(
                         row=row_index, column=0, sticky="w")
            tk.Entry(frame, textvariable=var, font=("Microsoft YaHei", 10)).grid(
                row=row_index, column=1, sticky="ew", padx=5, pady=2)
            row_index += 1
        existing = list(self.state.laws) + [""] * 3
        for i, value in enumerate(existing, start=1):
            var = tk.StringVar(value=value)
            self._s15_controls.append((f"法规条目{i}", var, "law"))
            tk.Label(frame, text=f"法规条目 {i}", bg="#FFFFFF", fg="#425466",
                     font=("Microsoft YaHei", 9), anchor="w").grid(row=row_index, column=0, sticky="w")
            tk.Entry(frame, textvariable=var, font=("Microsoft YaHei", 10)).grid(
                row=row_index, column=1, sticky="ew", padx=5, pady=2)
            row_index += 1
        frame.grid_columnconfigure(1, weight=1)

    # ---------- 状态、契约与产出 ----------
    def _collect_current(self) -> None:
        """只收集允许编辑控件；禁用控件不可能进入 write_items。"""
        for key, widget in self._widgets.items():
            value = widget.get("1.0", "end-1c").strip()
            old = self.state.values.get(key, "")
            if value != old:
                self.state.touched_sections.add(self.state.fields[key].section)
            if not value and old:
                self.state.cleared_keys.add(key)
            elif value:
                self.state.cleared_keys.discard(key)
            self.state.values[key] = value

        for key, controls in self._pictogram_vars.items():
            if key not in self._pictogram_touched:
                continue
            selected = ordered_codes(code for code, var in controls.items() if var.get())
            value = display_value(selected) if selected else "无"
            old = self.state.values.get(key, "")
            if value != old:
                self.state.touched_sections.add(2)
            if value == "无" and old not in {"", "无"}:
                self.state.cleared_keys.add(key)
            elif value != "无":
                self.state.cleared_keys.discard(key)
            self.state.values[key] = value

        if self._current_section == 9:
            for key, var in self._s9_alias_vars.items():
                selected = var.get().strip()
                if selected:
                    self.state.label_overrides[key] = selected

        if self._current_section == 3 and self._component_vars:
            comps = []
            for row in self._component_vars:
                item = {k: row[k].get().strip() for k in ("name", "cas", "conc")}
                if any(item.values()):
                    comps.append(item)
            if comps != self.state.components:
                self.state.touched_sections.add(3)
            self.state.components = comps
            product = next((x for x in self.fields_by_section[3] if x.label == "产品类型"), None)
            if product and product.key in self._widgets:
                self.state.product_type = self.state.values.get(product.key, "混合物") or "混合物"

        if self._current_section == 8 and self._bio_vars:
            rows = []
            for row in self._bio_vars:
                values = [v.get().strip() for v in row]
                if any(values):
                    rows.append(values)
            if rows != self.state.bio_rows:
                self.state.touched_sections.add(8)
            self.state.bio_rows = rows

        if self._current_section == 15 and self._s15_controls:
            specials = {
                label: var.get().strip()
                for label, var, semantic in self._s15_controls
                if semantic == "special"
            }
            if specials != self.state.s15_special:
                self.state.touched_sections.add(15)
            self.state.s15_special.update(specials)
            laws = [var.get().strip() for _, var, semantic in self._s15_controls if semantic == "law" and var.get().strip()]
            if laws != self.state.laws:
                self.state.touched_sections.add(15)
            self.state.laws = laws

    def _payload(self) -> dict:
        self._collect_current()
        payload = build_write_items(self.state)
        # GUI 输出的默认策略：未改字段保留模板，显式清空才写空值。
        payload["empty_policy"] = "overwrite" if self.state.cleared_keys else "preserve"
        payload["missing_policy"] = "preserve"
        validate_write_items_permissions(payload)
        return payload

    def _save_json(self) -> None:
        payload = self._payload()
        if not payload.get("sections"):
            messagebox.showwarning("没有修改", "当前没有检测到可写入字段变更。")
            return
        path = filedialog.asksaveasfilename(
            title="保存标准写入项 JSON", defaultextension=".json",
            initialfile=f"write_items_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
            filetypes=[("JSON", "*.json")])
        if not path:
            return
        Path(path).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        self.status_var.set(f"已保存写入项：{path}")
        messagebox.showinfo("保存成功", f"写入项已保存：\n{path}")

    def _overwrite(self) -> None:
        payload = self._payload()
        if not payload.get("sections"):
            messagebox.showwarning("没有修改", "当前没有检测到可写入字段变更。")
            return
        summary = "\n".join(f"S{k}: {len(v) if isinstance(v, list) else '产品类型/成分'} 项"
                            for k, v in payload["sections"].items())
        if not messagebox.askyesno("确认覆写", f"只会覆写固定模板白名单字段：\n\n{summary}\n\n确认继续？"):
            return
        default = f"{self.template_path.stem}_GUI覆写_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        out = filedialog.asksaveasfilename(
            title="产出 MSDS Word", defaultextension=".docx", initialfile=default,
            filetypes=[("Word 文档", "*.docx")])
        if not out:
            return
        try:
            overwrite_engine.overwrite(
                str(self.template_path), payload, str(out),
                sections={int(x) for x in payload["sections"]},
                empty_policy=payload.get("empty_policy", "preserve"),
                missing_policy="preserve", missing_text="")
            ok, problems = overwrite_engine.verify_output(
                str(self.template_path), str(out), payload,
                sections={int(x) for x in payload["sections"]})
            if not ok:
                raise RuntimeError("\n".join(problems))
        except Exception as exc:
            messagebox.showerror("覆写失败", str(exc))
            self.status_var.set("覆写失败，正式输出未确认")
            return
        self.status_var.set(f"产出成功：{out}")
        messagebox.showinfo("产出完成", f"Word 已产出并通过闭环校验：\n{out}")


__all__ = ["OverwriteFormWindow"]
