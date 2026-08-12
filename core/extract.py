# -*- coding: utf-8 -*-
"""MSDS 结构化检索与内容提取.

把读取结果 (ParseResult) 展开为分层的可检索条目:
    section (节) → 大标题 (节标题) → 小标题 (序号子标题/字段序号) → 字段+内容

输出层级文本与结构化 JSON, 支持:
  - 按节 / 大标题 / 小标题 / 字段名 / 内容关键字 检索
  - 批量处理多个 MSDS 文件, 统一提取指定字段
  - 导出为 TSV (表格) 或 JSON, 供下游批量清洗/覆写/入库
  - 标准 Excel 库表导出 (export_excel_table): 固定产出
    「第一行 Section / 第二行 序号+标签 / A列型号 / A1·A2留空」透视总表,
    内置规范: 页眉/页脚等 sub 不单独成列、剔除值全空父级/装饰标签、
    S9 同义标签归一化 + 空值标「无数据」、成分表多成分展开、序号取多数型号.

分层模型:
    ExtractedField.section    : 节号 0..16 (0=页眉页脚)
    ExtractedField.big_title  : 大标题 = 节标题 (如 "8.接触控制/个人防护")
    ExtractedField.sub_title  : 小标题 = 序号子标题或字段序号 (如 "8.1 暴露控制")
    ExtractedField.label      : 字段标签 (去序号, 如 "呼吸系统防护")
    ExtractedField.value      : 字段内容
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
from .structure import ParseResult, SectionData, SectionRow


@dataclass
class ExtractedField:
    """一个可检索的分层条目 (section → 大标题 → 小标题 → 字段+内容)."""
    section: int
    big_title: str = ""        # 大标题 (节标题)
    sub_title: str = ""        # 小标题 (序号子标题或字段序号前缀, 无则空)
    label: str = ""            # 字段标签 (去序号)
    value: str = ""            # 字段内容
    seq: str = ""              # 序号 (如 "8.1")
    kind: str = "field"        # "field" | "sub" | "note" | "component"
    editable: bool = True

    def full_label(self) -> str:
        """带序号前缀的完整标签 (如 '8.1 暴露控制' / '手部防护')."""
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


def search_fields(entries: Iterable[ExtractedField], query: str,
                  scope: str = "all") -> list[ExtractedField]:
    """按关键字检索. scope: 'all' | 'label' | 'value' | 'section'.

    - query 可含空格分隔多词 (AND 匹配)
    - 匹配范围: 标签/序号/标题/内容 (all) 或指定
    """
    terms = [_norm(t) for t in query.split() if t.strip()]
    if not terms:
        return list(entries)
    out = []
    for e in entries:
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
        else:  # note
            lines.append(f"{ind}· {e.value}")
    return "\n".join(lines)


def render_json(entries: Iterable[ExtractedField]) -> str:
    """JSON 输出 (含分层字段, 供程序消费)."""
    return json.dumps([e.to_dict() for e in entries], ensure_ascii=False, indent=2)


def render_tsv(entries: Iterable[ExtractedField], header: bool = True) -> str:
    """TSV 表格输出: 节 | 大标题 | 小标题 | 标签 | 内容."""
    rows = []
    if header:
        rows.append("节\t大标题\t小标题\t标签\t内容")
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
    label: str = ""            # 字段标签 (如 "呼吸系统防护")
    value: str = ""            # 字段内容
    kind: str = "field"        # "field" | "note" | "component"
    editable: bool = True
    index: int = 0             # 节内稳定序号 (供手动标注持久化)

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
      - 有 seq 的 field 行 (如 9.1 外观)   → 二级 BigTitleNode (kind=field, 带 value)
      - 无 seq 的 field 行                  → 三级 FieldNode, 归属当前二级下
      - note/成分                          → 三级 FieldNode
    与 GUI 表格 (序号|标题|内容) 完全对应.
    """
    nodes: list[SectionNode] = []
    for n in sorted(result.sections):
        sec = result.sections[n]
        sn = SectionNode(number=n, title=sec.title, full_title=sec.full_title,
                         is_component_table=sec.is_component_table)
        cur: BigTitleNode | None = None
        for row in sec.iter_rows():
            if row.kind == "section":
                continue
            if row.kind == "sub":
                cur = BigTitleNode(seq=row.seq, title=row.label, kind="sub",
                                   editable=row.editable, index=row.index)
                sn.big_titles.append(cur)
                continue
            if row.kind == "field":
                if row.seq:
                    # 有序号 → 二级 (如 9.1 外观: 乳白色液体)
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
            kept_children = [f for f in b.children if _field_hit(f, terms, scope)]
            if kept_children:
                kept_big.append(dataclasses.replace(b, children=kept_children))
        kept_direct = [f for f in sn.direct_fields if _field_hit(f, terms, scope)]
        if sn_hit:
            out.append(sn)          # 节命中 → 整节保留
        elif kept_big or kept_direct:
            out.append(dataclasses.replace(sn, big_titles=kept_big,
                                           direct_fields=kept_direct))
    return out


def _field_hit(f: FieldNode, terms: list[str], scope: str) -> bool:
    if scope == "label":
        return any(t in _norm(f.label) for t in terms)
    if scope == "value":
        return any(t in _norm(f.value) for t in terms)
    return any(t in _norm(f.label) + " " + _norm(f.value) for t in terms)


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
    if f.kind == "note":
        return f"· {f.value}"
    return f"{f.label}: {f.value}"


def render_tree_json(nodes: Iterable[SectionNode]) -> str:
    """三级树嵌套 JSON 输出."""
    return json.dumps([n.to_dict() for n in nodes], ensure_ascii=False, indent=2)
