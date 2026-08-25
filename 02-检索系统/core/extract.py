# -*- coding: utf-8 -*-
"""MSDS 结构化检索与内容提取.

把读取结果 (ParseResult) 展开为分层的可检索条目 (父子级):
    section (节) → 大标题 (节标题) → 小标题 (序号子标题/字段序号) → 标签 + 字段

术语 (检索系统约束规范):
    序号 = seq (如 11.5)   标签 = label (如 致突变性)
    字段 = value (内容文本)  父子级 = 节 → 大标题/小标题 → 标签 → 字段 的层级

输出层级文本与结构化 JSON, 支持:
  - 按节 / 大标题 / 小标题 / 标签 / 字段关键字 检索
  - 批量处理多个 MSDS 文件, 统一提取指定字段
  - 导出为 TSV (表格) 或 JSON, 供下游批量清洗/覆写/入库
  - 标准 Excel 库表导出 (export_excel_table): 固定产出
    「第一行 Section / 第二行 序号+标签 / A列型号 / A1·A2留空」透视总表,
    内置规范: 页眉/页脚等 sub 不单独成列、剔除字段全空父级/装饰标签、
    S9 同义标签归一化 + 空字段标「无数据」、成分表多成分展开、序号取多数型号.

分层模型:
    ExtractedField.section    : 节号 0..16 (0=页眉页脚)
    ExtractedField.big_title  : 大标题 = 节标题 (如 "8.接触控制/个人防护")
    ExtractedField.sub_title  : 小标题 = 序号子标题或字段序号 (如 "8.1 暴露控制")
    ExtractedField.label      : 标签 (去序号, 如 "呼吸系统防护")
    ExtractedField.value      : 字段 (内容文本)
"""
from __future__ import annotations

import dataclasses
import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

from .docx_reader import read_msds
from .pivot_table import build_pivot_table as export_excel_table
from .pivot_table import build_single_pivot as export_single_excel
from .structure import ParseResult, SectionData, SectionRow


@dataclass
class ExtractedField:
    """一个可检索的分层条目 (section → 大标题 → 小标题 → 标签 → 字段).

    术语 (检索系统约束规范): 序号 = seq, 标签 = label, 字段 = value.
    """
    section: int
    big_title: str = ""        # 大标题 (节标题)
    sub_title: str = ""        # 小标题 (序号子标题或字段序号前缀, 无则空)
    label: str = ""            # 标签 (去序号, 如 "致突变性")
    value: str = ""            # 字段 (内容文本)
    seq: str = ""              # 序号 (如 "11.5")
    kind: str = "field"        # "field" | "sub" | "note" | "component"
    editable: bool = True

    def full_label(self) -> str:
        """带序号前缀的完整标签 (如 '11.5 致突变性' / '手部防护')."""
        if self.seq and self.label:
            return f"{self.seq} {self.label}"
        return self.label or self.seq

    def to_dict(self) -> dict:
        return {
            "section": self.section,
            "big_title": self.big_title,
            "sub_title": self.sub_title,
            "label": self.label,
            "value": self.value,
            "seq": self.seq,
            "kind": self.kind,
            "editable": self.editable,
        }

    def to_row(self) -> list:
        """TSV 一行 (便于表格导入)."""
        return [str(self.section), self.big_title, self.sub_title,
                self.full_label(), self.value]


def _sub_title_of(row: SectionRow) -> str:
    """小标题: 序号子标题 → '8.1 暴露控制'; 有 seq 的字段 → 仅 seq 部分.

    归属规则: 字段行若自身带序号 (如 S9 的 9.1 外观), 小标题 = 该序号;
    否则沿用当前子标题上下文 (由调用方维护).
    """
    if row.kind == "sub":
        return f"{row.seq} {row.label}".strip() if row.seq else row.label
    return row.seq or ""


def _subtable_text(title: str, header: list[str],
                   rows: list[list[str]]) -> str:
    """子表 (S8.2 生物限值等) → 多行文本.

    每个数据行一行, 格式 "表头1: 值1 | 表头2: 值2 | ..."
    (空值跳过; 无表头时直接拼接各列值), 供扁平条目/检索使用.
    """
    lines: list[str] = []
    for r in rows:
        if header:
            pairs = [f"{h}: {v}" for h, v in zip(header, r) if str(v).strip()]
            lines.append(" | ".join(pairs) if pairs else " | ".join(r))
        else:
            lines.append(" | ".join(str(v) for v in r))
    return "\n".join(lines)


def iter_extracted(sec: SectionData) -> list[ExtractedField]:
    """把一个节的 iter_rows 展开为分层条目 (维护当前小标题上下文)."""
    out: list[ExtractedField] = []
    cur_sub = ""
    for row in sec.iter_rows():
        if row.kind == "section":
            continue
        if row.kind == "sub":
            cur_sub = _sub_title_of(row)
            out.append(ExtractedField(section=sec.number, big_title=sec.full_title,
                                      sub_title=cur_sub, label=row.label, value="",
                                      seq=row.seq, kind="sub", editable=False))
            continue
        if row.kind == "field":
            # 有 seq 的字段 (如 S9 9.1) → 其本身即小标题层级
            if row.seq:
                cur_sub = f"{row.seq} {row.label}".strip()
                out.append(ExtractedField(section=sec.number, big_title=sec.full_title,
                                          sub_title=cur_sub, label=row.label, value=row.value,
                                          seq=row.seq, kind="field", editable=row.editable))
            else:
                out.append(ExtractedField(section=sec.number, big_title=sec.full_title,
                                          sub_title=cur_sub, label=row.label, value=row.value,
                                          seq="", kind="field", editable=row.editable))
            continue
        if row.kind == "subtable":
            # 内嵌子表 (S8.2 生物限值等): 表头+数据行 → 多行文本条目,
            # 保证 "二甲苯" / "生物限值" 等可被检索且输出有实际内容
            out.append(ExtractedField(section=sec.number, big_title=sec.full_title,
                                      sub_title=cur_sub, label=row.label,
                                      value=_subtable_text(row.label,
                                                           row.sub_header,
                                                           row.sub_rows),
                                      seq="", kind="subtable", editable=False))
            continue
        # note: 通栏说明, 无标签
        out.append(ExtractedField(section=sec.number, big_title=sec.full_title,
                                  sub_title=cur_sub, label="", value=row.value,
                                  kind="note", editable=row.editable))
    return out


def extract_doc(result: ParseResult, include_component: bool = True) -> list[ExtractedField]:
    """把整个 ParseResult 展开为分层条目列表 (含 S3 成分表)."""
    out: list[ExtractedField] = []
    for n in sorted(result.sections):
        sec = result.sections[n]
        out.extend(iter_extracted(sec))
        if include_component and sec.is_component_table:
            # 成分表头 (实际识别到的表头, 如 'Chemical Name | CAS Number | %（w/w）')
            # 作为可检索条目 (与 GUI 表头一致); 未记录时用中文标准表头兜底
            hdr = sec.component_header or "化学品名称 | CAS编号 | 含量%（w/w）"
            out.append(ExtractedField(section=n, big_title=sec.full_title,
                                      sub_title="成分", label="成分表头",
                                      value=hdr,
                                      kind="component_header", editable=False))
            for c in sec.components:
                out.append(ExtractedField(section=n, big_title=sec.full_title,
                                          sub_title="成分", label=c.name,
                                          value=f"{c.name} | CAS: {c.cas} | 含量: {c.conc}",
                                          kind="component", editable=c.editable))
    return out


# ---- 图片导出 ----

def export_pictograms(result: ParseResult, out_dir: str | Path,
                      prefix: str = "") -> list[str]:
    """把 ParseResult.images (象形图原图) 导出为文件.

    命名: {prefix}{文件名}_{s节}_{序号}.{png|jpeg}
    返回导出文件路径列表 (与 images 顺序一一对应).
    """
    from .structure import ImageData
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = Path(result.file_name).stem or "doc"
    paths: list[str] = []
    for i, img in enumerate(result.images):
        fn = f"{prefix}{stem}_s{img.section}_{i}.{img.ext}"
        p = out_dir / fn
        p.write_bytes(img.blob)
        paths.append(str(p))
    return paths


# ---- 检索 ----

def _norm(s: str) -> str:
    """归一化: 去空白/冒号/中英文标点, 小写 (用于模糊匹配)."""
    return re.sub(r"[\s：:：,，。；;（）()\-]", "", (s or "")).lower()


def _is_page_number(e) -> bool:
    """是否 S0 页码字段 (如 "页码: 5 / 5").

    页码是页脚动态生成的页次信息, 无检索价值, 检索时一律排除
    (与导出侧 pivot_table/build_standard_table 的 _EXCLUDE 一致).
    """
    return e.section == 0 and (e.label or "").strip() == "页码"


def search_fields(entries: Iterable[ExtractedField], query: str,
                  scope: str = "all") -> list[ExtractedField]:
    """按关键字检索. scope: 'all' | 'label' | 'value' | 'section'.

    - query 可含空格分隔多词 (AND 匹配)
    - 匹配范围: 标签/序号/标题/内容 (all) 或指定
    - S0 页码字段不参与检索
    """
    terms = [_norm(t) for t in query.split() if t.strip()]
    if not terms:
        return [e for e in entries if not _is_page_number(e)]
    out = []
    for e in entries:
        if _is_page_number(e):
            continue
        if scope == "label":
            hay = _norm(e.full_label())
        elif scope == "value":
            hay = _norm(e.value)
        elif scope == "section":
            hay = _norm(str(e.section))
        else:  # "all": 标签 + 小标题 + 大标题 + 内容
            hay = (_norm(e.full_label()) + " " + _norm(e.sub_title) + " "
                   + _norm(e.big_title) + " " + _norm(e.value))
        if all(t in hay for t in terms):
            out.append(e)
    return out


def get_field(entries: Iterable[ExtractedField], section: int,
              label: str) -> ExtractedField | None:
    """精确定位: 第 section 节, 字段标签为 label (去序号), 返回首个匹配."""
    ln = _norm(label)
    for e in entries:
        if e.section == section and e.kind in ("field", "component") and _norm(e.label) == ln:
            return e
    return None


# ---- 输出 ----

def render_text(entries: Iterable[ExtractedField]) -> str:
    """分层文本输出: section → 大标题 → 小标题 → 字段: 内容."""
    lines: list[str] = []
    cur_section: int | None = None
    cur_big: str = ""
    for e in entries:
        # 节标题 + 大标题 (每次换节输出一次)
        if e.section != cur_section or e.big_title != cur_big:
            cur_section = e.section
            cur_big = e.big_title
            lines.append(f"[{e.section}] {e.big_title}")
        # 缩进层级: 有归属小标题 → 深一级 (8空格); 无 → 大标题下一级 (4空格)
        ind = "        " if e.sub_title else "    "
        if e.kind == "sub":
            lines.append(f"    ─ {e.full_label()}")
        elif e.kind == "field":
            lines.append(f"{ind}{e.full_label()}: {e.value}")
        elif e.kind == "component":
            # value 含 '名称 | CAS: .. | 含量: ..', 显示时去掉重复名称
            shown = e.value[len(e.label):].lstrip(" |")
            lines.append(f"{ind}∟ {e.label}  [{shown}]")
        elif e.kind == "component_header":
            lines.append(f"{ind}▤ {e.value}")
        elif e.kind == "subtable":
            # 内嵌子表 (S8.2 生物限值等): 标签行 + 每数据行 "表头: 值" 多行
            if e.label:
                lines.append(f"{ind}▤ {e.full_label()}")
            for line in (e.value or "").splitlines():
                lines.append(f"{ind}  · {line}")
        else:  # note
            lines.append(f"{ind}· {e.value}")
    return "\n".join(lines)


def render_json(entries: Iterable[ExtractedField]) -> str:
    """JSON 输出 (含分层字段, 供程序消费)."""
    return json.dumps([e.to_dict() for e in entries], ensure_ascii=False, indent=2)


def render_tsv(entries: Iterable[ExtractedField], header: bool = True) -> str:
    """TSV 表格输出: 节 | 大标题 | 小标题 | 标签 | 字段."""
    rows = []
    if header:
        rows.append("节\t大标题\t小标题\t标签\t字段")
    for e in entries:
        rows.append("\t".join(str(x).replace("\t", " ").replace("\n", " ") for x in e.to_row()))
    return "\n".join(rows)


def print_hierarchy(result: ParseResult, query: str | None = None,
                    scope: str = "all") -> int:
    """打印分层检索结果 (返回匹配条目数)."""
    entries = extract_doc(result)
    if query:
        entries = search_fields(entries, query, scope)
    print(f"文件: {result.file_name} | 共 {len(entries)} 条匹配")
    print(render_text(entries))
    return len(entries)


# ---- 批量提取 ----

def extract_many(paths: Iterable[str | Path], *,
                 query: str | None = None, scope: str = "all",
                 sections: set[int] | None = None) -> dict[str, list[ExtractedField]]:
    """批量处理多个 MSDS 文件, 返回 {文件名: 分层条目}.

    - query: 关键字过滤 (全部文件统一检索)
    - sections: 仅保留指定节号
    """
    result: dict[str, list[ExtractedField]] = {}
    for p in paths:
        try:
            r = read_msds(p)
        except Exception as exc:
            result[Path(p).name] = []
            print(f"⚠️ 读取失败 {p}: {exc}")
            continue
        entries = extract_doc(r)
        if sections:
            entries = [e for e in entries if e.section in sections]
        if query:
            entries = search_fields(entries, query, scope)
        result[Path(p).name] = entries
    return result


def export_tsv(result: ParseResult, path: str | Path, query: str | None = None,
               scope: str = "all") -> None:
    """导出分层检索结果为 TSV 文件."""
    entries = extract_doc(result)
    if query:
        entries = search_fields(entries, query, scope)
    Path(path).write_text(render_tsv(entries), encoding="utf-8-sig")


# ============================================================
# 三级父子级树模型 (对应 GUI 表格: 节 → 大标题/小标题 → 字段)
#   一级 SectionNode   : 节 (含 16 节标题)
#   二级 BigTitleNode  : 大标题/小标题 (序号子标题或有 seq 的字段标题)
#   三级 FieldNode     : 具体字段 (无 seq 的字段行, 归属到二级下)
# ============================================================


@dataclass
class FieldNode:
    """三级: 字段叶子 (归属到某二级大标题下)."""
    label: str = ""            # 标签 (如 "致突变性")
    value: str = ""            # 字段 (内容文本)
    kind: str = "field"        # "field" | "note" | "component" | "subtable"
    editable: bool = True
    index: int = 0             # 节内稳定序号 (供手动标注持久化)
    # kind == "subtable" 时携带内嵌子表 (如 S8.2 生物限值表):
    #   表头列 + 数据行 (每行 = 各列值), 供 GUI 以子表形式呈现
    sub_header: list[str] = field(default_factory=list)
    sub_rows: list[list[str]] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {"label": self.label, "value": self.value,
                "kind": self.kind, "editable": self.editable}


@dataclass
class BigTitleNode:
    """二级: 大标题/小标题 (序号子标题或有 seq 的字段行).

    自身可带 value (如有 seq 的字段行 "9.1 外观: 乳白色液体"),
    也可作为容器承载其下的无 seq 字段 (如 "8.1 暴露控制" 下挂 呼吸系统防护 等).
    """
    seq: str = ""
    title: str = ""
    value: str = ""
    kind: str = "sub"          # "sub" (纯子标题) | "field" (有seq字段行)
    editable: bool = True
    index: int = 0
    children: list[FieldNode] = field(default_factory=list)

    def full_title(self) -> str:
        if self.seq and self.title:
            return f"{self.seq} {self.title}"
        return self.title or self.seq

    def to_dict(self) -> dict:
        return {"seq": self.seq, "title": self.title, "value": self.value,
                "kind": self.kind, "editable": self.editable,
                "children": [c.to_dict() for c in self.children]}


@dataclass
class SectionNode:
    """一级: 节 (树根)."""
    number: int
    title: str = ""
    full_title: str = ""
    big_titles: list[BigTitleNode] = field(default_factory=list)  # 二级
    direct_fields: list[FieldNode] = field(default_factory=list)  # 无二级归属的三级字段
    is_component_table: bool = False

    def to_dict(self) -> dict:
        return {
            "section": self.number,
            "title": self.full_title,
            "big_titles": [b.to_dict() for b in self.big_titles],
            "fields": [f.to_dict() for f in self.direct_fields],
        }


def build_hierarchy(result: ParseResult, include_component: bool = True) -> list[SectionNode]:
    """把 ParseResult 构建为三级父子级树 (对应 GUI 表格内容).

    规则:
      - sub 行 (如 8.1 暴露控制)           → 二级 BigTitleNode (纯子标题)
      - 有 seq 的 field 行 (如 9.1 外观)   → 二级 BigTitleNode (kind=field, 带字段)
      - 无 seq 的 field 行                  → 三级 FieldNode, 归属当前二级下
      - note/成分                          → 三级 FieldNode
      - **S11 毒性资料特殊**: 字段行按国标大类 (GB/T 16483-2008 §11.1~11.10)
        归并 → 每个大类一个二级 BigTitleNode (seq=11.N, title=大类名,
        value=该大类所有子项归并值); 原始平铺子项 (LD50/物种/方法等) 不再
        逐行展开, 保证"检索页面从上到下 = 入库总表从左到右" (总表 S11 亦
        只产国标 11 大类别 + 总结句).
    与 GUI 表格 (序号|标签|字段) 完全对应.
    """
    from .s11 import S11_MAJOR_FIELDS, S11_MAJOR_SEQ, s11_group_rows
    nodes: list[SectionNode] = []
    for n in sorted(result.sections):
        sec = result.sections[n]
        sn = SectionNode(number=n, title=sec.title, full_title=sec.full_title,
                         is_component_table=sec.is_component_table)
        # S11: 先收集字段行 → 国标大类归并 (不逐个平铺)
        s11_groups: dict[str, list[str]] = {}
        if n == 11:
            s11_rows = [(row.label, row.value) for row in sec.iter_rows()
                        if row.kind == "field" and row.label.strip()]
            s11_groups = s11_group_rows(s11_rows) if s11_rows else {}
        # 共用结构原则 (2026-08-17 用户确认):
        #   所有带序号的行都是独立父级, 平级不归并嵌套 (S2 的 2.3/2.4/2.5
        #   不并入 2.2 标签要素; 四子列 2.6~2.9 各自独立) — 不再对任何节做
        #   标签要素父级化特殊处理, 统一走通用序号分派.
        rows_iter = sec.iter_rows()
        cur: BigTitleNode | None = None
        for row in rows_iter:
            if row.kind == "section":
                continue
            if row.kind == "subtable":
                # 内嵌子表已由下方 sub_tables 块单独挂载到父级, 跳过
                continue
            if n == 11 and (row.kind == "sub" or row.kind == "field"):
                # S11 子标题/字段已归并到国标大类, 不直接平铺
                continue
            if row.kind == "sub":
                cur = BigTitleNode(seq=row.seq, title=row.label, kind="sub",
                                   editable=row.editable, index=row.index)
                sn.big_titles.append(cur)
                continue
            if row.kind == "field":
                if row.seq:
                    # 有序号 → 独立父级二级 (带值; 如 9.1 外观 / 2.3 信号词)
                    cur = BigTitleNode(seq=row.seq, title=row.label, value=row.value,
                                       kind="field", editable=row.editable, index=row.index)
                    sn.big_titles.append(cur)
                elif cur is not None:
                    cur.children.append(FieldNode(label=row.label, value=row.value,
                                                  kind="field", editable=row.editable,
                                                  index=row.index))
                else:
                    sn.direct_fields.append(FieldNode(label=row.label, value=row.value,
                                                      kind="field", editable=row.editable,
                                                      index=row.index))
                continue
            # note
            fn = FieldNode(label="", value=row.value, kind="note",
                           editable=row.editable, index=row.index)
            if cur is not None:
                cur.children.append(fn)
            else:
                sn.direct_fields.append(fn)
        # 节内内嵌子表 (S8.2 生物限值等): 挂到对应父级二级下.
        # 父级由 title/seq 匹配 (如 "生物限值" / "8.2"); 无父级则独立为
        # 一个 sub 二级 (保持表结构可见). 子表节点作为三级呈现, 按 Word
        # 出现顺序 (sub_tables 追加顺序) 排列.
        for st in sec.sub_tables:
            parent = next((b for b in sn.big_titles
                           if b.title == st.title or b.seq == st.seq), None)
            if parent is None:
                parent = BigTitleNode(seq=st.seq, title=st.title, kind="sub",
                                      editable=False, index=len(sn.big_titles))
                sn.big_titles.append(parent)
            parent.children.append(FieldNode(
                label=st.title, value="", kind="subtable", editable=False,
                index=len(parent.children),
                sub_header=list(st.header), sub_rows=[list(r) for r in st.rows]))
        # S11 国标大类二级标题: 按国标固定顺序, 值 = 大类归并 (换行)
        if n == 11 and s11_groups:
            for major in S11_MAJOR_FIELDS:
                vals = s11_groups.get(major)
                if not vals:
                    continue
                sn.big_titles.append(BigTitleNode(
                    seq=S11_MAJOR_SEQ[major], title=major, value="\n".join(vals),
                    kind="field", editable=True, index=len(sn.big_titles)))
        if include_component and sec.is_component_table:
            for ci, c in enumerate(sec.components):
                sn.direct_fields.append(FieldNode(
                    label=c.name, value=f"{c.name} | CAS: {c.cas} | 含量: {c.conc}",
                    kind="component", editable=c.editable, index=ci))
        nodes.append(sn)
    return nodes


def flatten_nodes(nodes: Iterable[SectionNode]) -> list[ExtractedField]:
    """把三级树扁平化为 ExtractedField 列表 (兼容旧 API/TSV)."""
    out: list[ExtractedField] = []
    for sn in nodes:
        for b in sn.big_titles:
            if b.kind == "sub":
                out.append(ExtractedField(section=sn.number, big_title=sn.full_title,
                                          sub_title=b.full_title(), label=b.title,
                                          value="", seq=b.seq, kind="sub", editable=b.editable))
            elif b.value:
                out.append(ExtractedField(section=sn.number, big_title=sn.full_title,
                                          sub_title=b.full_title(), label=b.title,
                                          value=b.value, seq=b.seq, kind="field",
                                          editable=b.editable))
            for f in b.children:
                if f.kind == "subtable":
                    # 内嵌子表 → 单条目, label=子表标题, value=表头:值 多行文本
                    # (检索/矩阵/TSV 均能看到子表内容, 而非空 label/value)
                    out.append(ExtractedField(section=sn.number, big_title=sn.full_title,
                                              sub_title=b.full_title(),
                                              label=f.label or "子表",
                                              value=_subtable_text(f.label,
                                                                   f.sub_header,
                                                                   f.sub_rows),
                                              seq="", kind="subtable",
                                              editable=False))
                    continue
                out.append(ExtractedField(section=sn.number, big_title=sn.full_title,
                                          sub_title=b.full_title(), label=f.label,
                                          value=f.value, kind=f.kind, editable=f.editable))
        for f in sn.direct_fields:
            out.append(ExtractedField(section=sn.number, big_title=sn.full_title,
                                      sub_title="", label=f.label, value=f.value,
                                      kind=f.kind, editable=f.editable))
    return out


def search_tree(nodes: Iterable[SectionNode], query: str,
                scope: str = "all") -> list[SectionNode]:
    """在三级树上检索, 保留父子关系 (命中字段时连同其节/大标题一起保留)."""
    terms = [_norm(t) for t in query.split() if t.strip()]
    if not terms:
        return list(nodes)
    out: list[SectionNode] = []
    for sn in nodes:
        # 节级命中: section 检索按节号; all/label 检索中节标题不触发整节保留
        # (label 检索应精确到字段, 否则 "供应商" 命中节标题会把整节带出)
        sn_hit = False
        if scope == "section":
            sn_hit = any(t in _norm(str(sn.number)) for t in terms)
        elif scope == "all":
            sn_hit = any(t in _norm(sn.full_title) for t in terms)
        kept_big: list[BigTitleNode] = []
        kept_direct: list[FieldNode] = []
        for b in sn.big_titles:
            b_text = _norm(b.full_title()) + " " + _norm(b.value)
            b_hit = any(t in b_text for t in terms)
            if scope == "label":
                b_hit = any(t in _norm(b.full_title()) for t in terms)
            elif scope == "value":
                b_hit = any(t in _norm(b.value) for t in terms)
            if b_hit:
                kept_big.append(b)   # 二级命中 → 保留整棵二级子树
                continue
            # 三级字段命中 → 保留该字段, 连带二级容器
            kept_children = [f for f in b.children
                             if _field_hit(f, terms, scope, sn.number)]
            if kept_children:
                kept_big.append(dataclasses.replace(b, children=kept_children))
        kept_direct = [f for f in sn.direct_fields
                       if _field_hit(f, terms, scope, sn.number)]
        if sn_hit:
            out.append(sn)          # 节命中 → 整节保留
        elif kept_big or kept_direct:
            out.append(dataclasses.replace(sn, big_titles=kept_big,
                                           direct_fields=kept_direct))
    return out


def _field_hit(f: FieldNode, terms: list[str], scope: str,
               section: int | None = None) -> bool:
    """三级字段是否命中检索词.

    section: 所属节号 (用于排除 S0 页码等无检索价值字段, 与扁平检索一致).
    """
    if section == 0 and (f.label or "").strip() == "页码":
        return False
    if scope == "label":
        return any(t in _norm(f.label) for t in terms)
    if scope == "value":
        return any(t in _norm(f.value) for t in terms)
    text = _norm(f.label) + " " + _norm(f.value)
    if f.kind == "subtable":
        # 内嵌子表 (S8.2 生物限值): 表头列 + 数据行 全部参与检索,
        # 使 "二甲苯" / "ACGIH" / "生物限值" 等都能命中该子表
        text += " " + " ".join(_norm(h) for h in f.sub_header)
        text += " " + " ".join(_norm(c) for r in f.sub_rows for c in r)
    return any(t in text for t in terms)


def render_tree(nodes: Iterable[SectionNode]) -> str:
    """三级父子级树形文本输出 (对应 GUI 表格内容)."""
    lines: list[str] = []
    for sn in nodes:
        lines.append(f"[{sn.number}] {sn.full_title}")
        for b in sn.big_titles:
            if b.kind == "field" and b.value:
                lines.append(f"    ├─ {b.full_title()}: {b.value}")
            else:
                lines.append(f"    ├─ {b.full_title()}")
            for ci, f in enumerate(b.children):
                end = "└─" if ci == len(b.children) - 1 else "├─"
                lines.append(f"    │   {end} {_render_leaf(f)}")
        for f in sn.direct_fields:
            lines.append(f"    ├─ {_render_leaf(f)}")
    return "\n".join(lines)


def _render_leaf(f: FieldNode) -> str:
    """三级叶子的显示文本 (component 去重复名称)."""
    if f.kind == "component":
        shown = f.value[len(f.label):].lstrip(" |") if f.value.startswith(f.label) else f.value
        return f"∟ {f.label}  [{shown}]"
    if f.kind == "subtable":
        return (f"▤ {f.label}  [{len(f.sub_rows)}行 × {len(f.sub_header)}列]"
                if f.label else f"▤ 子表  [{len(f.sub_rows)}行 × {len(f.sub_header)}列]")
    if f.kind == "note":
        return f"· {f.value}"
    return f"{f.label}: {f.value}"


def render_tree_json(nodes: Iterable[SectionNode]) -> str:
    """三级树嵌套 JSON 输出."""
    return json.dumps([n.to_dict() for n in nodes], ensure_ascii=False, indent=2)
