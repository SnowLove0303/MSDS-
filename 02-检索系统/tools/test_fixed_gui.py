# -*- coding: utf-8 -*-
"""测试修复后的 GUI 目录树与节点标题."""

import re

SEC_TITLES: dict[int, str] = {
    0: "0. 页眉/页脚", 1: "1. 物料及供应商标识", 2: "2. 危险性概述",
    3: "3. 成分/组成资料", 4: "4. 急救措施", 5: "5. 消防措施",
    6: "6. 意外泄漏措施", 7: "7. 操作和储存", 8: "8. 接触控制/个人防护",
    9: "9. 物理和化学特性", 10: "10. 稳定性和反应性", 11: "11. 毒性资料",
    12: "12. 生态信息", 13: "13. 处理注意事项", 14: "14. 运输信息",
    15: "15. 法规信息", 16: "16. 其他信息",
}

SEC_PURE_TITLES: dict[int, str] = {
    0: "页眉/页脚", 1: "物料及供应商标识", 2: "危险性概述",
    3: "成分/组成资料", 4: "急救措施", 5: "消防措施",
    6: "意外泄漏措施", 7: "操作和储存", 8: "接触控制/个人防护",
    9: "物理和化学特性", 10: "稳定性和反应性", 11: "毒性资料",
    12: "生态信息", 13: "处理注意事项", 14: "运输信息",
    15: "法规信息", 16: "其他信息",
}

def format_tree_title(num: int, title: str) -> str:
    t = (title or "").strip()
    t_clean = re.sub(r"^(?:第\s*\d+\s*节\s*[\.、：:\s]*|\d+[\.、\s]+)", "", t).strip()
    return f"{num}. {t_clean}"

print("=== 修复后目录树标题展示 ===")
for num in range(17):
    # 模拟各种来源的 title
    t_from_db = SEC_TITLES[num]
    t_from_pure = SEC_PURE_TITLES[num]
    print(f"Section {num:<2} -> '{format_tree_title(num, t_from_db)}'")
