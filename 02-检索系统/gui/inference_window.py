# -*- coding: utf-8 -*-
"""Section 2 推断工作台。

推断工作台不重复实现表单：左侧直接嵌入 ``S139FormPanel``，该面板同时被
独立的 17 节表单窗口复用；右侧直接使用现有 ``SectionView`` 表格组件。
"""
from __future__ import annotations

import sys
import tkinter as tk
from pathlib import Path
from tkinter import messagebox

from core.structure import SectionRow
from gui.theme import (
    COLOR_ACCENT,
    COLOR_BG,
    COLOR_BORDER,
    COLOR_NAV,
    COLOR_ORANGE,
    COLOR_PANEL,
    COLOR_TEXT,
)

from .form_window import S139FormPanel
from .section_tree import SectionView


ENGINE_PROGRAM_DIR = Path(r"F:\正式项目与模块化内容\冠志\MSDS\04-推断引擎\推断引擎程序")
if str(ENGINE_PROGRAM_DIR) not in sys.path:
    sys.path.insert(0, str(ENGINE_PROGRAM_DIR))

from s2_inference import infer_s2  # noqa: E402


_RESULT_SECTIONS = {"2": "S2 危险性概述"}
_NO_DATA = "无数据"
_S2_FIELDS = (
    ("2.1", "GHS危险性类别"),
    ("2.2", "象形图"),
    ("2.3", "信号词"),
    ("2.4", "危险性说明"),
    ("2.5", "防范说明"),
    ("2.6", "物理和化学危险"),
    ("2.7", "健康危害"),
    ("2.8", "环境危害"),
    ("2.9", "其他危害"),
)


class InferenceWindow(tk.Toplevel):
    """共用 S1/S3/S9 表单 → Section 2 推导结果。"""

    def __init__(self, master):
        super().__init__(master)
        self.title("推断引擎 · S1/S3/S9 → Section 2")
        self.geometry("1460x860")
        self.minsize(1180, 700)
        self.configure(bg=COLOR_BG)
        self._inference: dict | None = None
        self._result_status_var = tk.StringVar(value="尚未推导")

        self._build_layout()

    def _build_layout(self) -> None:
        header = tk.Frame(self, bg=COLOR_NAV, padx=14, pady=10)
        header.pack(fill="x")
        tk.Label(header, text="Section 推断", bg=COLOR_NAV, fg="#FFFFFF",
                 font=("Microsoft YaHei", 15, "bold")).pack(side="left")
        tk.Label(header, text="使用已有 S1/S3/S9 表单，生成可审计的 Section 2 候选结果",
                 bg=COLOR_NAV, fg="#D9E5F5",
                 font=("Microsoft YaHei", 10)).pack(side="left", padx=(16, 0))
        tk.Label(header, textvariable=self._result_status_var,
                 bg=COLOR_NAV, fg="#FFFFFF",
                 font=("Microsoft YaHei", 10, "bold")).pack(side="right")

        body = tk.Frame(self, bg=COLOR_BG)
        body.pack(fill="both", expand=True, padx=8, pady=8)

        left = tk.Frame(body, bg=COLOR_PANEL, highlightthickness=1,
                        highlightbackground=COLOR_BORDER, width=680)
        left.pack(side="left", fill="both", expand=True, padx=(0, 6))
        left.pack_propagate(False)
        right = tk.Frame(body, bg=COLOR_PANEL, highlightthickness=1,
                         highlightbackground=COLOR_BORDER, width=680)
        right.pack(side="left", fill="both", expand=True, padx=(6, 0))
        right.pack_propagate(False)

        self._build_input_panel(left)
        self._build_result_panel(right)

    def _build_input_panel(self, parent: tk.Frame) -> None:
        top = tk.Frame(parent, bg=COLOR_PANEL, padx=10, pady=8)
        top.pack(fill="x")
        tk.Label(top, text="输入事实", bg=COLOR_PANEL, fg=COLOR_TEXT,
                 font=("Microsoft YaHei", 12, "bold")).pack(side="left")
        tk.Label(top, text="已接入表单系统", bg=COLOR_PANEL, fg=COLOR_ORANGE,
                 font=("Microsoft YaHei", 9)).pack(side="left", padx=(12, 0))
        tk.Button(top, text="推导", command=self._run_inference,
                  bg=COLOR_ACCENT, fg="#FFFFFF", relief="flat", cursor="hand2",
                  font=("Microsoft YaHei", 10, "bold"), padx=18,
                  pady=4).pack(side="right")
        tk.Button(top, text="清空输入", command=self._clear_inputs,
                  bg="#E8EAED", fg=COLOR_TEXT, relief="flat", cursor="hand2",
                  font=("Microsoft YaHei", 9), padx=10,
                  pady=4).pack(side="right", padx=(0, 6))

        # 唯一输入来源：复用独立表单窗口中的 S139FormPanel。
        self._form_panel = S139FormPanel(parent, show_actions=False)
        self._form_panel.pack(fill="both", expand=True, padx=8, pady=(0, 8))

    def _build_result_panel(self, parent: tk.Frame) -> None:
        top = tk.Frame(parent, bg=COLOR_PANEL, padx=10, pady=8)
        top.pack(fill="x")
        tk.Label(top, text="推导结果", bg=COLOR_PANEL, fg=COLOR_TEXT,
                 font=("Microsoft YaHei", 12, "bold")).pack(side="left")
        tk.Label(top, text="沿用 GUI SectionView 表格 · 仅候选不写库",
                 bg=COLOR_PANEL, fg=COLOR_ORANGE,
                 font=("Microsoft YaHei", 9)).pack(side="right")

        tk.Label(parent, text="Section 目录", bg=COLOR_PANEL, fg="#80868B",
                 font=("Microsoft YaHei", 9, "bold"), anchor="w",
                 padx=10).pack(fill="x")
        self._result_list = tk.Listbox(
            parent, height=3, exportselection=False, bg="#F7F9FC",
            fg=COLOR_TEXT, selectbackground=COLOR_ACCENT,
            selectforeground="#FFFFFF", relief="solid", bd=1,
            font=("Microsoft YaHei", 10), activestyle="none",
        )
        self._result_list.pack(fill="x", padx=10, pady=(3, 8))
        self._result_list.bind("<<ListboxSelect>>", self._on_result_section)

        # 与导入检索、数据库检索完全共用同一个表格组件。
        self._result_view = SectionView(parent)
        self._result_view.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self._render_empty_result()

    # ---------- 表单 → 推断事实包 ----------

    def _build_package(self) -> dict:
        form = self._form_panel.get_data()
        s1_data = form["s1"]
        s9_rows = form["s9"]
        usages = []
        for name, cas, concentration in form["s3_components"]:
            if not any((name, cas, concentration)):
                continue
            usages.append({
                "raw_name": name or _NO_DATA,
                "cas_no": cas or _NO_DATA,
                "raw_cas": cas or _NO_DATA,
                "concentration": concentration or _NO_DATA,
            })
        return {
            "package_version": "gui-0.2.0",
            "engine": "section2",
            "model": s1_data.get("产品名称") or "GUI表单输入",
            "input_snapshot": {
                "s1": {"facts": {k: v or _NO_DATA for k, v in s1_data.items()}, "rows": []},
                "s3": {"facts": {"产品类型": form["s3_product_type"] or _NO_DATA},
                        "rows": usages},
                "s9": {
                    "facts": {row["label"]: row["value"] or _NO_DATA
                              for row in s9_rows if row.get("label")},
                    "rows": [],
                },
            },
            "substance_snapshot": {
                "usage": usages,
                "substance_count": len(usages),
                "usage_count": len(usages),
            },
        }

    def _run_inference(self) -> None:
        try:
            self._inference = infer_s2(self._build_package())
        except Exception as exc:
            messagebox.showerror("推导失败", str(exc))
            return
        self._result_list.delete(0, "end")
        for key, label in _RESULT_SECTIONS.items():
            self._result_list.insert("end", f"{key}  {label}")
        if self._result_list.size():
            self._result_list.selection_set(0)
            self._result_list.activate(0)
        status = self._inference.get("status", "manual_review")
        self._result_status_var.set(f"S2：{status}")
        self._render_result_section("2")

    # ---------- SectionView 表格结果 ----------

    def _on_result_section(self, _event=None) -> None:
        selection = self._result_list.curselection()
        if not selection:
            return
        self._render_result_section(list(_RESULT_SECTIONS)[selection[0]])

    def _render_empty_result(self) -> None:
        rows = [
            SectionRow(kind="section", label="S2 危险性概述", value="",
                       editable=False, span=True),
            SectionRow(kind="note", label="提示",
                       value="请在左侧表单填写 S1 / S3 / S9 后点击“推导”",
                       editable=False, span=True),
        ]
        self._result_view.show_rows(2, "S2 危险性概述", rows,
                                    meta="尚未生成推导结果")

    def _render_result_section(self, section: str) -> None:
        if section != "2" or not self._inference:
            self._render_empty_result()
            return
        inference = self._inference
        result = inference.get("result") or {}
        status = inference.get("status", "manual_review")
        rows: list[SectionRow] = [
            SectionRow(kind="section", label="S2 危险性概述 · 推导结果",
                       value="", editable=False, span=True),
            SectionRow(kind="field", seq="", label="推导状态", value=status,
                       editable=False, index=0),
            SectionRow(kind="sub", label="危险性概述", value="",
                       editable=False),
        ]
        for index, (seq, label) in enumerate(_S2_FIELDS, start=1):
            rows.append(SectionRow(
                kind="field", seq=seq, label=label,
                value=str(result.get(label) or _NO_DATA),
                editable=False, index=index,
            ))

        conditions = inference.get("conditions") or []
        if conditions:
            rows.append(SectionRow(kind="section", label="判断条件",
                                   value="", editable=False, span=True))
            for condition in conditions:
                matched = "命中" if condition.get("matched") else "未命中"
                inputs = condition.get("inputs") or {}
                input_text = "；".join(f"{key}={value}" for key, value in inputs.items())
                basis = "、".join(str(x) for x in condition.get("basis") or [])
                rows.append(SectionRow(
                    kind="field", seq="", label=str(condition.get("id", "rule")),
                    value=(f"状态：{matched}\n输入：{input_text}\n依据：{basis}\n"
                           f"结果：{condition.get('result', _NO_DATA)}"),
                    editable=False, index=len(rows),
                ))

        reasons = inference.get("review_reasons") or []
        if reasons:
            rows.append(SectionRow(kind="section", label="人工复核原因",
                                   value="", editable=False, span=True))
            for index, reason in enumerate(reasons, start=1):
                rows.append(SectionRow(
                    kind="field", seq="", label=f"复核原因 {index}",
                    value=str(reason), editable=False, index=len(rows),
                ))

        self._result_view.show_rows(
            2, "S2 危险性概述 · 推导结果", rows,
            meta=f"状态：{status} · 结果为候选值，仅供审核，不自动写入正式库",
        )

    def _clear_inputs(self) -> None:
        self._form_panel.reset_data()
        self._inference = None
        self._result_status_var.set("尚未推导")
        self._result_list.delete(0, "end")
        self._render_empty_result()


__all__ = ["InferenceWindow"]
