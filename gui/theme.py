# -*- coding: utf-8 -*-
"""GUI 主题: 深色导航 + 浅色内容, 与国彩/冠志文档风格统一."""
from __future__ import annotations

# 配色
COLOR_BG = "#F5F6FA"            # 主背景
COLOR_PANEL = "#FFFFFF"         # 面板
COLOR_NAV = "#1F2A44"           # 左侧导航深蓝
COLOR_NAV_SEL = "#2E3D5C"       # 导航选中
COLOR_ACCENT = "#1B6FD0"        # 强调蓝
COLOR_GREEN = "#1E8E3E"         # 成功
COLOR_ORANGE = "#E8710A"        # 警告
COLOR_RED = "#C5221F"           # 错误
COLOR_TEXT = "#202124"          # 正文
COLOR_GRAY = "#80868B"          # 次要文字
COLOR_BORDER = "#DADCE0"        # 边框
COLOR_ROW_ALT = "#F1F3F4"       # 表格交替行
COLOR_YELLOW = "#FCE8B2"        # 残留风险高亮

# 覆写指向 → 颜色/说明
WRITE_COLORS = {
    "template": COLOR_GRAY,     # 保留模板
    "product": COLOR_GREEN,     # 产品覆盖
    "clear": COLOR_GRAY,        # 清空
    "add": COLOR_ACCENT,        # 新增
    "review": COLOR_ORANGE,     # 需人工确认
}
WRITE_LABELS = {
    "template": "保留模板",
    "product": "产品覆盖",
    "clear": "清空",
    "add": "新增",
    "review": "需确认",
}
