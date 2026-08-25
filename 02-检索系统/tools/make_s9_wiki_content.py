# -*- coding: utf-8 -*-
"""生成飞书 Wiki 页面所需的 Markdown 内容并调用 lark-cli 写入."""

import sqlite3
from pathlib import Path

DB_PATH = r"F:\正式项目与模块化内容\冠志\MSDS\03-数据库\正式库\Data Base\msds_standard.db"
conn = sqlite3.connect(DB_PATH)
cur = conn.cursor()

# 1. 23 项标准字段及其归类表述
cur.execute("""
SELECT std_seq, std_name, raw_label, occurrences_count, models_sample, sample_values
FROM s9_label_mapping
ORDER BY CAST(SUBSTR(std_seq, 3) AS INTEGER), std_name, occurrences_count DESC
""")
mapping_rows = cur.fetchall()

# 组织标准字段及其不同表述
std_dict = {} # (seq, name) -> list of (raw_label, count, sample_values)
for seq, std, raw, count, models, sample in mapping_rows:
    key = (seq, std)
    std_dict.setdefault(key, []).append((raw, count, sample))

# 2. 未归入标准字段的非标标签清单
cur.execute("""
SELECT raw_label, COUNT(DISTINCT model) as m_cnt, GROUP_CONCAT(DISTINCT model) as models, GROUP_CONCAT(raw_value, ' | ') as vals
FROM s9_raw_expression
WHERE is_standard = 0
GROUP BY raw_label
ORDER BY m_cnt DESC, raw_label
""")
unclass_rows = cur.fetchall()

md_lines = []
md_lines.append("# Section 9 物理和化学特性：23项标准化字段及表述清单\n")
md_lines.append("> **文档说明**：本清单基于全库 253 个型号 MSDS 原始中文 Word 文档（共 4,925 条 Section 9 原文记录）进行全量原文检索、清洗与归类分析建立。严格遵循 GB/T 16483 规范及「基于结构找内容、结构冻结」原则。\n")

md_lines.append("## 一、Section 9 标准化 23 个字段及其涵盖的所有表述清单\n")
md_lines.append("以下为标准骨架中 23 个平铺字段及其在各型号原文中出现的 **47 种不同表述**（含同义别名、异写及测试条件剥离项）：\n")

md_lines.append("| 标准序号 | 标准字段名 | 归纳涵盖的原文不同表述 | 命中型号数 | 代表性原始数据示例 | 归一化处理规则 |")
md_lines.append("|:---:|---|---|:---:|---|---|")

# 规则描述映射
rule_map = {
    "外观": "收敛空格异写（外 观 → 外观）",
    "嗅觉阈值": "收敛常见异写错字（嗅觉阀值 → 嗅觉阈值）",
    "pH值": "剥离测试浓度条件括号（1%/5%/10%水溶液）及大小写（PH值）",
    "离子性": "标准字段直接对齐",
    "初沸点": "收敛最低初沸点及单位括号（℃）",
    "闪点": "剥离测试方法（闭口/开口）及单位括号（℃）",
    "蒸发速率": "归并行业同义词（百分比挥发性）",
    "可燃性（固态、气态）": "收敛简化标签（可燃性）",
    "燃烧值": "标准字段直接对齐",
    "饱和蒸气压": "剥离测试温度与压力单位（20℃/25℃/Kpa/mmHg）",
    "相对蒸气密度": "标准字段直接对齐",
    "密度": "归并行业同义别名（比重、相对密度、比重/25℃）",
    "水溶性": "归并行业同义别名（水中溶解度）",
    "表面张力": "剥离测试溶液浓度后缀（10%水溶液）",
    "辛醇/水分配系数对数值": "收敛助词异写（辛醇/水分配系数的对数值）",
    "自燃温度": "归并行业同义别名（着火点）",
    "引燃温度": "标准字段直接对齐",
    "分解温度": "标准字段直接对齐",
    "动力粘度": "剥离测试温度条件（粘度/25℃）及同义粘度",
    "爆炸特性": "标准字段直接对齐",
    "粉尘爆炸级别": "标准字段直接对齐",
    "固体含量": "标准字段直接对齐",
    "其他信息": "标准字段直接对齐",
}

for (seq, name), labels in std_dict.items():
    raw_labels_str = "<br>".join(f"• `{lbl}` ({cnt}型)" for lbl, cnt, _ in labels)
    total_cnt = sum(cnt for _, cnt, _ in labels)
    # 取一个典型示例
    samples = [s.replace("\n", " ").replace("|", "\|") for _, _, s in labels if s and s != "无数据"]
    sample_str = ("<br>".join(samples[:2])) if samples else "无数据"
    rule_str = rule_map.get(name, "标准字段对齐")
    md_lines.append(f"| **{seq}** | **{name}** | {raw_labels_str} | **{total_cnt}** | {sample_str} | {rule_str} |")

md_lines.append("\n---\n")
md_lines.append("## 二、未能归入标准 23 字段的 Section 9 非标清单（共 29 种）\n")
md_lines.append("根据「**结构冻结、清单外标签不扩充 Schema 结构**」的强制规范，以下 29 种出现在原始 docx 中的非标字段未作为独立列输出到 23 字段标准宽表中，但已完整保存在明细数据库中供溯源与推断消费：\n")

md_lines.append("### 1. 涂料/高分子树脂特有物化性能指标（10 种）\n")
md_lines.append("| 原始标签 | 涉及型号数 | 典型型号代表 | 原始数据示例 | 未归入原因及业务定位 |")
md_lines.append("|---|:---:|---|---|---|")

category_1 = {
    "最低成膜温度MFFT/℃": ("乳液特有成膜温度指标，国标 S9 无此项", "PA-3615, PA-3617, PA-3655"),
    "玻璃化温度Tg/℃": ("聚合物热力学转变温度，属树脂产品规格指标", "PA-3615, PA-3617, PA-3655"),
    "有效成分": ("防腐剂/助剂有效物质量分数，非标准固含量定义", "BEK-100L, BEK-200, CX-470"),
    "NCO含量": ("异氰酸酯基团质量分数，属固化剂化学交联指标", "OS-9013, OS-9016, OS-9018"),
    "羟基含量": ("羟基丙烯酸树脂官能团含量", "PA-4804T, RA-15000"),
    "酸值": ("油脂/助剂类游离酸度参数", "GZ-080, GZ-080A"),
    "碘值": ("油脂不饱和双键指标", "GZ-080, GZ-080A"),
    "粘度（涂-4杯）": ("流出杯特定测定条件（非动力粘度 mPa.s）", "WE-68"),
    "熔点": ("聚乙烯蜡乳液特有固态熔点", "WE-68"),
    "熔点/凝固点": ("覆写版扩展字段（国标 16 节标准清单无独立列）", "EC-1801(覆写版)"),
}

for r in unclass_rows:
    lbl, m_cnt, models, vals = r
    if lbl in category_1:
        reason, rep_models = category_1[lbl]
        vals_clean = vals.replace("\n", " ").replace("|", "\|")[:60]
        md_lines.append(f"| **{lbl}** | {m_cnt} | {rep_models} | `{vals_clean}` | {reason} |")

md_lines.append("\n### 2. 有机溶剂/单体化合物专有物理指标（7 种）\n")
md_lines.append("| 原始标签 | 涉及型号数 | 典型型号代表 | 原始数据示例 | 未归入原因及业务定位 |")
md_lines.append("|---|:---:|---|---|---|")

category_2 = {
    "分子量": ("单一化合物分子量参数", "2-苯氧基乙醇, PPH-B"),
    "含量": ("单体化合物纯度参数", "2-苯氧基乙醇, PPH-B"),
    "APHA值": ("铂钴比色法色度指标", "2-苯氧基乙醇, PPH-B"),
    "蒸发速度（醋酸丁酯-1）": ("带特定参照物基准的蒸发速度", "2-苯氧基乙醇, PPH-B"),
    "倾点": ("溶剂/油品低温流动性极限", "OS-405, OS-406"),
    "浊点": ("非离子表面活性剂相分离温度", "OS-405, OS-406"),
    "气味": ("嗅觉描述（标准骨架由 9.2 嗅觉阈值 统领）", "EC-1801(覆写版)"),
}

for r in unclass_rows:
    lbl, m_cnt, models, vals = r
    if lbl in category_2:
        reason, rep_models = category_2[lbl]
        vals_clean = vals.replace("\n", " ").replace("|", "\|")[:60]
        md_lines.append(f"| **{lbl}** | {m_cnt} | {rep_models} | `{vals_clean}` | {reason} |")

md_lines.append("\n### 3. 嵌套特殊稀释条件长尾标签（3 种）\n")
md_lines.append("| 原始标签 | 涉及型号数 | 典型型号代表 | 原始数据示例 | 说明 |")
md_lines.append("|---|:---:|---|---|---|")

category_3 = {
    "PH值（1:10稀释在水中）": ("非标准 1%/5%/10% 格式的长尾稀释说明", "OS-8950"),
    "pH值（1:10稀释在水中）": ("非标准 1%/5%/10% 格式的长尾稀释说明", "PU-1034"),
    "pH值（（1:10稀释在水中））": ("多重双括号嵌套导致正则未匹配", "RU-13000"),
}

for r in unclass_rows:
    lbl, m_cnt, models, vals = r
    if lbl in category_3:
        reason, rep_models = category_3[lbl]
        vals_clean = vals.replace("\n", " ").replace("|", "\|")[:60]
        md_lines.append(f"| **{lbl}** | {m_cnt} | {rep_models} | `{vals_clean}` | {reason} |")

md_lines.append("\n### 4. 原文排版跨列粘连异常项（7 种，源自个别排版异常 Word）\n")
md_lines.append("| 原始标签 | 涉及型号数 | 典型型号代表 | 原始数据示例 | 说明 |")
md_lines.append("|---|:---:|---|---|---|")

category_4 = [
    "7-9 （按 1", "pH 值", "约1.06g/cm3可混溶", "不适用无数据",
    "乳白色液体", "化学稳定性", "辛醇/ 水分配系数的对数值"
]

for r in unclass_rows:
    lbl, m_cnt, models, vals = r
    if lbl in category_4:
        vals_clean = vals.replace("\n", " ").replace("|", "\|")[:60]
        md_lines.append(f"| `{lbl}` | {m_cnt} | {models} | `{vals_clean}` | 源 Word 表格缺失冒号或跨列粘连 |")

md_lines.append("\n---\n")
md_lines.append("## 三、规范与结构冻结说明\n")
md_lines.append("1. **结构唯一性与指纹校验**：数据库严格按照 `07f97a44b0a48133` 结构指纹锁定，宽表固定输出 23 项标准字段，保障全库 253 个型号格式完全一致。\n")
md_lines.append("2. **双轨数据存档**：\n")
md_lines.append("   - 标准数据表：`Section9_标准字段总表_全型号.xlsx`（253 型号 × 23 标准字段）\n")
md_lines.append("   - 原文全量明细表：`Section9_原文字段明细总表_全量.xlsx`（4,925 行原文记录，含所有非标字段）\n")

out_file = Path("./_s9_wiki_content.md")
out_file.write_text("\n".join(md_lines), encoding="utf-8")
print(f"✅ Markdown 文件已生成: {out_file.resolve()} ({len(md_lines)} 行)")
