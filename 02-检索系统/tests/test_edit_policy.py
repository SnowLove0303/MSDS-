# -*- coding: utf-8 -*-
"""固定模板 GUI 编辑权限测试。"""
from types import SimpleNamespace

import pytest

from core.edit_policy import (
    is_editable_field,
    is_editable_note,
    is_editable_row,
    validate_write_items_permissions,
)


def _row(kind, label, value=""):
    return SimpleNamespace(kind=kind, label=label, value=value)


def test_s1_parent_and_allowed_field_permissions():
    assert not is_editable_field(1, "1.1 产品名称")
    assert is_editable_field(1, "中文名称")
    assert is_editable_field(1, "1.2 产品使用建议和使用限制")


def test_s0_allowed_fields_and_page_lock():
    for label in ("Version", "产品名称", "公司名称", "产品型号", "修订日期"):
        assert is_editable_field(0, label)
    assert not is_editable_field(0, "页码")


def test_s9_extension_fields_are_editable():
    for label in ("有效成分", "玻璃化温度", "最低成膜温度", "NCO含量", "羟基含量",
                  "熔点/凝固点", "酸值", "碘值", "倾点", "浊点", "分子量",
                  "含量", "APHA值", "HLB值"):
        assert is_editable_field(9, label)


def test_protected_s8_parent_and_note_policy():
    assert not is_editable_field(8, "8.1 暴露控制")
    assert not is_editable_field(8, "8.2 生物限值")
    assert not is_editable_field(11, "", kind="note")
    assert is_editable_note(11, value="该产品无可用的毒理学研究。")
    assert not is_editable_note(11, value="未授权的通栏文本")
    assert not is_editable_note(4, value="任何通栏文本")


def test_row_kind_is_part_of_permission():
    assert is_editable_row(1, _row("field", "中文名称"))
    assert not is_editable_row(1, _row("sub", "产品名称"))
    assert not is_editable_row(3, _row("subtable", "成分"))


def test_write_items_reject_non_whitelist_field():
    valid = {"sections": {"1": [{"label": "中文名称", "value": "新值"}]}}
    validate_write_items_permissions(valid)
    with pytest.raises(ValueError, match="白名单"):
        validate_write_items_permissions({
            "sections": {"1": [{"label": "供应商信息", "value": "越权"}]}
        })


def test_write_items_allow_s8_bio_subtable_only():
    validate_write_items_permissions({
        "sections": {"8": [{"label": "生物限值", "subtable": {"rows": []}}]}
    })
    with pytest.raises(ValueError):
        validate_write_items_permissions({
            "sections": {"8": [{"label": "暴露控制", "value": "越权"}]}
        })
