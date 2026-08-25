# -*- coding: utf-8 -*-
"""
clean_write_items.py — 从 write_items_*.json 剔除"仅说明(note)"内容，生成纯 MSDS 版写入项。

背景：推导方案 content 区分「可写入 MSDS 的正文（msds）」与「仅说明（note：推导理由/
编辑指令/字段提醒）」。覆写引擎应只把 msds 内容写进标准表格模板，故在覆写前先清洗写入项。

规则与 msds-inference-write 技能的 content 区分（rules/output_structure.md）保持一致：
- 删除整段说明（段首"说明：本品…"/"推导要点："）；
- 删除括号内编辑指令（如"（建议设置专用公共邮箱）"）；
- 个别括号保留对 MSDS 有用的部分（如"（如铝、镁粉）""（低于-5℃有冻结风险）"）。

用法:
    python clean_write_items.py <write_items_in.json> [<write_items_pure.json>]
    默认输出 <同名>_pure.json

流水线位置：
    write_items_*.json → clean_write_items.py → write_items_*_pure.json
        → msds_overwrite_engine.py --write-items 纯版 → 标准化输出（纯 MSDS，无说明）
"""
import json, io, os, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# note 括号：整串删除
DEL_PARENS = [
    "（建议设置专用公共邮箱）",
    "（24小时，建议签约委托）",
    "（如产品需库房管理时填写）",
    "（N,N-二甲基乙醇胺含量下整体不达严重眼损伤类别，但仍应防止入眼）",
    "（如无中国限值，注明“未建立”，可参考ACGIH/其他来源或据实际监测确定）",
    "（如无中国限值，注明“未建立”）",
    "（接近水的沸点）",
    "（以水为主，蒸发缓慢）",
    "（金属盐类可能存在个体接触性致敏风险，建议皮肤接触者注意）",
    "（需在合规焚烧设施中进行）",
]
# note 括号：替换为保留的 MSDS 有用信息
REPL_PARENS = [
    ("（如铝、镁粉——含水的乳液与活泼金属可能反应）", "（如铝、镁粉）"),
    ("（以水为连续相，接近水的凝固点；低于-5℃有冻结风险）", "（低于-5℃有冻结风险）"),
]
# note 段落：以这些前缀开头的整段删除（按 \n 分段的段）
DEL_PARAGRAPH_PREFIX = ["说明：本品各健康危害成分", "推导要点：", "【说明】"]

def clean_text(value):
    if not value:
        return value
    paras = value.split("\n")
    out = []
    for p in paras:
        s = p.strip()
        if any(s.startswith(pre) for pre in DEL_PARAGRAPH_PREFIX):
            continue
        for d in DEL_PARENS:
            s = s.replace(d, "")
        for old, new in REPL_PARENS:
            s = s.replace(old, new)
        s = s.strip()
        if s:
            out.append(s)
    # 合并连续空行为单段
    return "\n".join(out)

def clean_items(items, path):
    n_hit = 0
    if isinstance(items, list):
        for it in items:
            if "value" in it:
                v = it["value"]
                nv = clean_text(v)
                if nv != v:
                    it["value"] = nv
                    n_hit += 1
    elif isinstance(items, dict):
        for k, v in items.items():
            if isinstance(v, str):
                nv = clean_text(v)
                if nv != v:
                    items[k] = nv
                    n_hit += 1
    return n_hit

def main():
    src = sys.argv[1] if len(sys.argv) > 1 else "write_items_os1330.json"
    dst = sys.argv[2] if len(sys.argv) > 2 else src.replace(".json", "_pure.json")
    data = json.load(open(src, encoding="utf-8"))
    total = 0
    for sec, items in data["sections"].items():
        total += clean_items(items, src)
    json.dump(data, open(dst, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    print(f"[OK] {os.path.basename(src)} → {os.path.basename(dst)}，清洗命中 {total} 个字段")

if __name__ == "__main__":
    main()
