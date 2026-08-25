# -*- coding: utf-8 -*-
"""MSDS 固定模板的字段编辑权限策略。

权限以已确认的 17 节父子级覆写边界为准，而不是直接信任 Word 解析器对
``editable`` 的启发式结果。节标题、父级结构、子表表头和页码域永远不可编辑；
只有白名单字段、动态子表数据行和规定的通栏说明槽位允许进入覆写表单。
"""
from __future__ import annotations

import re
from typing import Any

from .schema import standard_name


# 固定模板中允许编辑的标准字段。未列出的 field 一律不可编辑。
EDITABLE_FIELDS: dict[int, set[str]] = {
    0: {"Version", "产品名称", "公司名称", "产品型号", "修订日期"},
    1: {"中文名称", "Product name", "化学品分类", "产品使用建议和使用限制", "供应商名称",
        "供应商地址", "电话", "传真"},
    2: {"GHS危险性类别", "GHS标签要素", "GHS象形图", "信号词", "危险性说明",
        "防范说明", "物理和化学危险", "健康危害", "环境危害", "其他危害"},
    3: {"产品类型"},
    4: {"一般措施", "误服", "接触眼睛", "接触皮肤", "吸入"},
    5: {"合适的灭火剂", "不合适的灭火剂", "物质或混合物的特殊危害",
        "消防预防措施和保护设备"},
    6: {"个人预防措施、应急程序", "环境保护措施", "污染物收集和清除的方法"},
    7: {"安全操作防范", "安全储存条件"},
    8: {"呼吸系统防护", "手部防护", "防护手套的合适材料", "氟化橡胶 –FKM",
        "丁基橡胶 –IIR", "丁腈橡胶 – NBR", "建议", "眼睛防护", "身体防护", "工程控制"},
    9: {"外观", "嗅觉阈值", "pH值", "离子性", "初沸点", "闪点",
        "蒸发速率", "可燃性（固态、气态）", "燃烧值", "饱和蒸气压", "相对蒸气密度",
        "密度", "水溶性", "表面张力", "辛醇/水分配系数对数值", "自燃温度",
        "引燃温度", "分解温度", "动力粘度", "爆炸特性", "粉尘爆炸级别",
        "固体含量", "有效成分", "玻璃化温度", "最低成膜温度", "NCO含量",
        "羟基含量", "熔点/凝固点", "酸值", "碘值", "倾点", "浊点", "分子量",
        "含量", "APHA值", "HLB值", "其他信息"},
    10: {"化学稳定性", "危险分解产物", "可能的危害反应", "应避免的条件", "禁配物"},
    11: {"急性毒性", "主要皮肤刺激性", "主要眼睛刺激性", "致敏性", "致突变性",
         "致癌性", "生殖毒性", "特异性靶器官系统毒性（一次接触/反复接触）",
         "吸入危险", "附加信息"},
    12: {"生态毒性", "持久性和降解性", "其他不利的影响"},
    13: {"处理方法"},
    14: {"公路和铁路运输", "海上运输", "空运", "用户特殊注意事项"},
}

# 这些节允许编辑经过语义识别的通栏说明槽位；其它 note 不进入表单写入。
EDITABLE_NOTE_SECTIONS = {11, 12, 13, 15, 16}

# 正式 Word/PDF 输出禁止把这些状态占位词当作业务事实呈现。内部输入单、
# 检索结果和审计日志仍可保留原始状态；这里提供共享判定，供表单契约和
# 覆写引擎在进入正式容器前做最后一道过滤。
FORMAL_PLACEHOLDER_VALUES = frozenset({
    "无数据", "无数据资料", "无可用数据", "无资料", "未提供", "不适用",
    "no data", "no data available", "not applicable", "not available",
})


def is_formal_placeholder(value: Any) -> bool:
    """判断一个值是否仅为禁止呈现的缺失/不适用占位词。"""
    text = re.sub(r"\s+", " ", str(value or "").strip()).casefold()
    return text in {re.sub(r"\s+", " ", item).casefold()
                    for item in FORMAL_PLACEHOLDER_VALUES}

# S9 模板可能没有物理行、但规范允许按需克隆参考行插入的扩充字段。
DYNAMIC_FIELD_SEQS: dict[int, dict[str, str]] = {
    9: {
        "有效成分": "9.23", "玻璃化温度": "9.24", "最低成膜温度": "9.25",
        "NCO含量": "9.26", "羟基含量": "9.27", "熔点/凝固点": "9.28",
        "酸值": "9.29", "碘值": "9.30", "倾点": "9.31", "浊点": "9.32",
        "分子量": "9.33", "含量": "9.34", "APHA值": "9.35", "HLB值": "9.36",
    },
}

# 固定模板父级/结构标签，禁止作为字段写入。
PROTECTED_LABELS: dict[int, set[str]] = {
    0: {"页眉", "页脚"},
    1: {"产品名称", "供应商信息"},
    2: {"紧急情况概述"},
    3: {"成分"},
    8: {"暴露控制", "生物限值"},
}


def normalize_edit_label(label: Any) -> str:
    """归一化权限比较用标签。"""
    value = str(label or "")
    value = re.sub(r"^\d+(?:\.\d+)*[.．、\s　]*", "", value)
    value = re.sub(r"\s+", "", value).replace("：", ":").rstrip(":")
    return value


def is_editable_field(section: int, label: str, *, kind: str = "field") -> bool:
    """判断固定模板中的普通 field 是否允许编辑。"""
    normalized = normalize_edit_label(label)
    # S9 先用原始表述命中 Schema aliases，再做去空格权限比较；如果先归一化
    # 掉英文别名中的空格，会导致 pH value(1% aqueous solution) 无法命中。
    canonical = normalize_edit_label(standard_name(section, label)) if section == 9 else normalized
    if normalized in {normalize_edit_label(x) for x in PROTECTED_LABELS.get(section, set())}:
        return False
    if kind == "note":
        return False
    allowed = {normalize_edit_label(x) for x in EDITABLE_FIELDS.get(section, set())}
    return normalized in allowed or canonical in allowed


def is_editable_note(section: int, label: str = "", value: str = "") -> bool:
    """按通栏说明槽位语义判断 note 是否允许编辑。"""
    if section not in EDITABLE_NOTE_SECTIONS:
        return False
    text = str(value or label or "").strip()
    if section in {11, 12}:
        return text.startswith("该产品无可用") or text.startswith("以下为")
    if section == 13:
        return (
            text.startswith("必须遵守适用的国标")
            or text.startswith("在欧盟领域内废弃")
            or text == "欧盟废弃细则"
        )
    if section == 15:
        # 飞书《覆写规范》边界表：S15 的指引段、"其它的规定"、
        # "符合下列法规要求"均为锁定结构；只有其下方法规条目由专用
        # 动态控件管理。普通 note 不能因为有文本而被误判为可编辑。
        return False
    if section == 16:
        return text.startswith("就我们所掌握的知识信息") or text == "免责声明"
    return False


def is_editable_row(section: int, row: Any) -> bool:
    """按固定模板权限判断一个解析行。"""
    kind = getattr(row, "kind", "")
    if kind == "note":
        return is_editable_note(section, getattr(row, "label", ""), getattr(row, "value", ""))
    if kind != "field":
        return False
    return is_editable_field(section, getattr(row, "label", ""), kind=kind)


def is_editable_component(section: int, component: Any) -> bool:
    """判断动态成分子表数据行权限。"""
    return section == 3 and bool(getattr(component, "editable", True))


def is_editable_subtable_data(section: int, label: str) -> bool:
    """判断动态子表数据行权限；表头永远不可编辑。"""
    return section in {3, 8} and normalize_edit_label(label) in {"成分", "生物限值"}


def _is_allowed_note_item(section: int, label: str) -> bool:
    """判断覆写契约中的通栏说明语义键。"""
    value = str(label or "").strip()
    if section in {11, 12}:
        return value in {"产品说明", "成分参考引导段"}
    if section == 13:
        return (
            value.startswith("必须遵守适用的国标")
            or value.startswith("在欧盟领域内废弃")
            or value == "欧盟废弃细则"
        )
    if section == 15:
        # S15 结构标题/指引段的正文由专用控件管理；契约允许保留两个
        # 明确命名的输入槽位，避免 GUI 与 Web GUI 因白名单不同步而丢字段。
        return (
            value.startswith("法规条目")
            or value in {"其它的规定", "符合下列法规要求"}
        )
    if section == 16:
        return value.startswith("就我们所掌握的知识信息") or value == "免责声明"
    return False


def validate_write_items_permissions(write_items: dict) -> None:
    """拒绝 GUI/自动化写入项中的非白名单字段。"""
    if not isinstance(write_items, dict):
        raise ValueError("write_items 必须是对象")
    sections = write_items.get("sections", {})
    if not isinstance(sections, dict):
        raise ValueError("write_items.sections 必须是对象")
    for raw_sec, payload in sections.items():
        try:
            sec = int(raw_sec)
        except (TypeError, ValueError) as exc:
            raise ValueError(f"非法节号: {raw_sec}") from exc
        if sec < 0 or sec > 16:
            raise ValueError(f"节号超出 S0~S16: {sec}")
        if sec == 3 and isinstance(payload, dict):
            allowed_keys = {"产品类型", "components"}
            if set(payload) - allowed_keys:
                raise ValueError("S3 只允许产品类型和 components")
            components = payload.get("components", [])
            if not isinstance(components, list):
                raise ValueError("S3 components 必须是列表")
            for component in components:
                if not isinstance(component, dict) or set(component) - {"name", "cas", "conc"}:
                    raise ValueError("S3 成分只允许 name/cas/conc")
            continue
        if isinstance(payload, list):
            for item in payload:
                if not isinstance(item, dict):
                    raise ValueError(f"S{sec} 写入项必须是对象")
                allowed_item_keys = {"seq", "label", "value", "subtable"}
                if set(item) - allowed_item_keys:
                    raise ValueError(f"S{sec} 写入项含未授权键")
                label = str(item.get("label", ""))
                if "subtable" in item:
                    if not (sec == 8 and normalize_edit_label(label) == "生物限值"):
                        raise ValueError(f"S{sec} 子表标签不在允许范围: {label}")
                    subtable = item["subtable"]
                    if not isinstance(subtable, dict) or set(subtable) - {"header", "rows"}:
                        raise ValueError("S8 生物限值子表格式非法")
                    rows = subtable.get("rows", [])
                    if not isinstance(rows, list) or any(not isinstance(row, list) or len(row) != 5 for row in rows):
                        raise ValueError("S8 生物限值必须是五列数据行")
                    continue
                if is_editable_field(sec, label, kind="field") or _is_allowed_note_item(sec, label):
                    continue
                raise ValueError(f"S{sec} 字段不在固定模板编辑白名单: {label}")
        elif isinstance(payload, dict):
            raise ValueError(f"S{sec} 不允许使用字典字段写入格式")


__all__ = [
    "DYNAMIC_FIELD_SEQS", "EDITABLE_FIELDS", "EDITABLE_NOTE_SECTIONS", "PROTECTED_LABELS",
    "FORMAL_PLACEHOLDER_VALUES", "is_editable_component", "is_editable_field",
    "is_editable_note", "is_editable_row", "is_formal_placeholder",
    "is_editable_subtable_data", "normalize_edit_label", "validate_write_items_permissions",
]
