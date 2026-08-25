# -*- coding: utf-8 -*-
"""
build_initial_content.py — 从 facts.json 自动生成 16 节完整的 content.json 初稿。

作用：保证每次使用技能产出的方案"16 节齐全、字段完整"，再由 AI 依据联网检索逐节完善。
初稿的每节 content[] 含：
  - 从原 MSDS 提取的既定事实（S1/S3/S9 自动带入）
  - 依据 rules/hazards_lib.md 自动识别已收录成分
  - 其余需推导的节用"待推导"占位，并按 output_structure.md 列出必填字段提示

用法:
    python build_initial_content.py "<facts.json>" --product "产品名" --out "<content.json>"
    # 可选: --company "供应商" --date "2026-08-14"
"""
import argparse, io, json, os, re, sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 16 节标准标题（GB/T 16483）
SECTION_TITLES = [
    "化学品及企业标识", "危险性概述", "成分/组成信息", "急救措施", "消防措施",
    "泄漏应急处理", "操作处置与储存", "接触控制和个体防护", "理化特性",
    "稳定性和反应性", "毒理学信息", "生态学信息", "废弃处置", "运输信息",
    "法规信息", "其他信息",
]

# 每节必填字段提示（对照 rules/output_structure.md）
SECTION_FIELDS = {
    1: ["中英文名", "商品名", "产品类型", "推荐+限制用途", "供应商全信息(名称/地址/邮编/电话/传真/邮箱/24h应急)", "SDS编号/版本/日期"],
    2: ["紧急情况概述", "GHS危险性类别", "标签要素(象形图/信号词/危险性说明/防范说明)", "物理和化学危险", "健康危害(吸入/皮肤/眼睛/食入/慢性)", "环境危害", "其他危害"],
    3: ["产品类型=混合物", "组分按含量递减", "每项=名称+CAS+含量+危害简述", "保密组分注明商业机密但列危险性"],
    4: ["吸入/皮肤/眼睛/食入", "眼冲>=15min", "禁止催吐条件", "症状和影响", "对施救者忠告", "对医生提示"],
    5: ["适用灭火剂", "不适用灭火剂", "特别危险性(含燃烧产物)", "灭火注意事项(SCBA/消防废水)"],
    6: ["人员防护(具体装备)", "环境保护措施", "收容清除方法及材料", "防止次生灾害"],
    7: ["操作(通风/防静电/防火/禁配物/卫生)", "储存(温度范围/远离火种热源/禁配隔离/储区应急)"],
    8: ["职业接触限值(GBZ 2.1,含皮/敏/Gx)", "生物限值", "监测方法", "工程控制", "个体防护(呼吸/手/眼/身体)"],
    9: ["外观性状", "气味", "pH", "熔点凝固点", "沸点沸程", "闪点", "爆炸极限", "蒸气压", "蒸气密度", "相对密度", "溶解性", "分配系数", "自燃温度", "分解温度", "其他(粘度/固含等)"],
    10: ["稳定性", "危险反应", "应避免的条件", "禁配物", "危险的分解产物"],
    11: ["概述(未整体试验按组分)", "急性毒性", "皮肤刺激腐蚀", "眼睛刺激腐蚀", "过敏", "致突变", "致癌", "生殖", "STOT一次", "STOT反复", "吸入危害", "附加信息"],
    12: ["生态毒性(LC50/EC50+物种+方法)", "持久性和降解性", "生物累积性", "土壤迁移性", "PBT/vPvB", "其他不良影响"],
    13: ["废弃处置方法", "污染包装物", "废弃注意事项(引中国法规)"],
    14: ["UN号", "运输名称", "危险类别", "包装组", "海洋污染物", "各运输方式", "特殊注意事项(非危险品填不适用)"],
    15: ["法律法规清单(现行版)", "目录状态", "组分法规信息"],
    16: ["编制依据", "版本修订记录", "缩略语", "培训建议", "参考文献", "免责声明"],
}

def load_hazards_lib():
    """解析 rules/hazards_lib.md 里 '### 化学名（CAS xxx）' 小节名，用于成分匹配。"""
    path = os.path.join(SKILL_DIR, "rules", "hazards_lib.md")
    entries = []
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                m = re.match(r"^###\s+(.+?)（CAS\s*([0-9-]+)）", line.strip())
                if m:
                    entries.append({"name": m.group(1), "cas": m.group(2)})
    return entries

def load_limits():
    """读取浓度限值表含 ≥ 的行（文本），供初稿引用。"""
    path = os.path.join(SKILL_DIR, "rules", "limits.md")
    lines = []
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                s = line.strip()
                if s.startswith("|") and "≥" in s:
                    lines.append(s.strip("|").strip())
    return lines

_S3_HEAD_WORDS = ("化学品名称", "CAS编号", "%（w/w）", "w/w", "成分")

def parse_s3(facts):
    """从 facts.json 的 s3_raw 尽力提取 名称/CAS/含量。

    以"条目"为单位处理（不 join、不拆行），兼容常见模板：
      A) 三列多行：一条内 '名称块 | CAS块 | 含量块'，块内每行一个，zip 对齐；
      B) 每行一条：'名称 CAS 含量'。
    """
    items = facts.get("s3_raw", [])
    names, cas_list, pcts = [], [], []

    # 模板 A：含 | 的条目
    for item in items:
        if "|" not in item:
            continue
        if item.count("|") >= 2 and any(w in item for w in _S3_HEAD_WORDS):
            continue  # 表头行
        parts = [p.strip() for p in item.split("|")]
        if len(parts) < 2:
            continue
        blk = [[x.strip() for x in p.split("\n") if x.strip()] for p in parts]
        if len(blk) >= 3:
            nb, cb, pb = blk[0], blk[1], blk[2]
        else:
            # 2 段：第二段可能是 CAS（数字-数字）或含量（%/＞/＜等）
            if any(re.search(r"\d{2,7}-\d{2}-\d{1}", x) for x in blk[1]):
                nb, cb, pb = blk[0], blk[1], []
            else:
                nb, cb, pb = blk[0], [], blk[1]
        if nb and nb[0].startswith("产品类型"):
            continue
        maxlen = max(len(nb), len(cb), len(pb))
        for i in range(maxlen):
            nm = nb[i] if i < len(nb) else ""
            ca = cb[i] if i < len(cb) else ""
            pc = pb[i] if i < len(pb) else ""
            if nm:
                names.append(nm); cas_list.append(ca); pcts.append(pc)
    if names:
        return list(zip(names, cas_list, pcts))

    # 模板 B：每行 名称 空格 CAS 空格 含量
    raw_lines = [x.strip() for x in "\n".join(items).split("\n") if x.strip()]
    for ln in raw_lines:
        m = re.search(r"([A-Za-z0-9一-鿿（）()\- ]+?)[ \t]+(\d{2,7}-\d{2}-\d{1})(?:[ \t]+(.+))?", ln)
        if m:
            names.append(m.group(1).strip().strip("：: |"))
            cas_list.append(m.group(2))
            pcts.append((m.group(3) or "").strip())
    return list(zip(names, cas_list, pcts))

def build_diagnosis_from_fields(facts):
    """根据缺失节与常见铁律生成诊断汇总（占位，AI 精修）。"""
    rows = []
    missing = facts.get("missing_sections", [])
    if missing:
        rows.append(["整体", "缺少节: " + ", ".join("S%d" % n for n in missing), "高"])
    rows += [
        ["1 物料及供应商标识", "检查是否缺英文名/邮编/邮箱/24h应急电话/SDS编号/最初编制日期", "中"],
        ["2 危险性概述", "检查是否缺紧急情况概述/信号词/危害分述/标签要素完整", "高"],
        ["8 接触控制/个人防护", "职业接触限值须引中国GBZ 2.1，不得写无可用的接触限值信息", "高"],
        ["11 毒性资料", "严禁混入与本品无关成分的毒理数据", "高"],
        ["13 废弃处置", "须引中国固废法/危险废物名录，不用欧盟EWC", "中"],
        ["15 法规信息", "法规用现行版（危险化学品条例→国务院令591号）；补GB 30000系列/GB 15258", "高"],
    ]
    return rows

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("facts_json")
    ap.add_argument("--product", required=True, help="产品名（用于封面与文件命名）")
    ap.add_argument("--company", default="", help="供应商名称")
    ap.add_argument("--date", default="2026-08-14", help="推导日期")
    ap.add_argument("--out", required=True, help="content.json 输出路径")
    args = ap.parse_args()

    with open(args.facts_json, "r", encoding="utf-8") as f:
        facts = json.load(f)

    haz = load_hazards_lib()
    limits = load_limits()
    comps = parse_s3(facts)
    # 成分正文（可写入 MSDS）与速查库提示（仅说明，给 AI 用）分开
    comp_lines, comp_notes = [], []
    if comps:
        for name, cas, pct in comps:
            found = next((h for h in haz if h["cas"] == cas), None)
            comp_lines.append("%s（%s）%s" % (name, cas, pct))
            comp_notes.append("速查库%s，AI应结合联网检索完善该成分危害" %
                              ("已收录，请完善危害表述" if found else "未见收录，需联网检索"))
    else:
        comp_lines = []
        comp_notes = ["成分需从原MSDS第3部分核对并列出（名称/CAS/含量，按含量递减）"]

    s1_lines = facts.get("s1_raw", [])[:14]
    s9_lines = facts.get("s9_raw", [])[:30]

    def md(text, note=False):
        """content 条目：note=True 仅说明（不进纯 MSDS），否则为可写 MSDS 正文。"""
        return {"t": "note", "text": text} if note else {"t": "msds", "text": text}

    sections = []
    for n, title in enumerate(SECTION_TITLES, start=1):
        content = []
        if n == 1:
            content = [md(x) for x in s1_lines] + [
                md("【需补充】英文名称/邮编/邮箱/24h应急电话/SDS编号/最初编制日期/推荐与限制用途", True)]
        elif n == 3:
            content = [md(x) for x in comp_lines] + [md(x) for x in comp_notes] + [
                md("【需补充】产品类型=混合物；保密组分注明商业机密但列危险性", True)]
        elif n == 9:
            content = [md(x) for x in s9_lines] + [
                md("【需补充】气味/熔点凝固点/爆炸极限/蒸气压/自燃温度/分解温度（无数据写无数据/不适用）", True)]
        elif n == 2:
            content = [md("【待推导】紧急情况概述 / GHS危险性类别 / 标签要素 / 物理化学危险 / 健康危害 / 环境危害 / 其他危害", True)]
        elif n == 8:
            content = [md("【待推导】职业接触限值（GBZ 2.1，含皮/敏/Gx）+ 工程控制 + 个体防护（呼吸/手/眼/身体）", True)]
        elif n == 11:
            content = [md("【待推导】按各危险组分逐项：急性/刺激/致敏/致突变/致癌/生殖/STOT一次/STOT反复/吸入危害", True)]
        else:
            content = [md("【待推导】", True)]
        content.append(md("【必填字段】" + "；".join(SECTION_FIELDS[n]), True))
        sections.append({"num": "第%d部分" % n, "title": title,
                         "diagnosis": "（待精修）见诊断汇总表及铁律表",
                         "content": content})

    content = {
        "title": "%s 化学品安全技术说明书（MSDS）规范化推导方案" % args.product,
        "sds": {
            "product": args.product,
            "company": args.company,
            "date": args.date,
            "sds_no": "",
            "version": "",
            "basis": "GB/T 16483-2008、GB/T 17519-2013"
        },
        "meta": [
            "推导对象：%s%s" % (args.product, ("（" + args.company + "）") if args.company else ""),
            "推导依据：GB/T 16483-2008、GB/T 17519-2013、GB 30000 系列、GBZ 2.1-2019 及成分公开毒理数据库",
            "推导日期：" + args.date,
            "说明：本初稿由 build_initial_content.py 自动生成骨架，须由 AI 依据联网检索逐节完善后再渲染 Word。"
        ],
        "overview": [
            "本方案对原 MSDS（%s）进行规范性诊断，并依据第1部分（物料及供应商标识）、第3部分（成分/组成信息）、第9部分（物理和化学特性）的既定事实，反向推导其余各部分（第2、4~8、10~16）应当如何编写。" % os.path.basename(facts.get("source", "")),
            "推导总逻辑：产品身份 + 已知成分 + 已知理化性质 => 各成分 GHS 危害分类 => 混合物整体分类与标签要素 => 急救、消防、泄漏、操作储存、接触控制、稳定反应、毒理、生态、废弃、运输、法规各部分的规范表述。",
            "推导判据（详见 rules/limits.md）：急性毒性/皮肤腐蚀/眼刺激/STOT/水环境 ≥1.0%；过敏/致癌/生殖 ≥0.1%；吸入危害 ≥10% 且黏度条件。"
        ],
        "diagnosis_header": ["部分", "原文存在问题", "严重程度"],
        "diagnosis_rows": build_diagnosis_from_fields(facts),
        "basis_blocks": [
            {"title": "2.1 标准与规范依据", "items": [
                "GB/T 16483-2008《化学品安全技术说明书 内容和项目顺序》",
                "GB/T 17519-2013《化学品安全技术说明书 编写指南》",
                "GB 30000.1~29-2013《化学品分类和标签规范》系列",
                "GBZ 2.1-2019《工作场所有害因素职业接触限值》",
                "《危险化学品目录》（2015版，2022年调整）",
            ]},
            {"title": "2.2 混合物危害组分浓度限值（关键判据）", "items": limits},
            {"title": "2.4 成分依据（第3部分，从原MSDS提取）", "items": comp_lines + comp_notes},
            {"title": "2.5 理化性质依据（第9部分，从原MSDS提取）", "items": s9_lines},
        ],
        "sections": sections,
        "references": [
            "GB/T 16483-2008《化学品安全技术说明书 内容和项目顺序》",
            "GB/T 17519-2013《化学品安全技术说明书 编写指南》",
            "GBZ 2.1-2019《工作场所有害因素职业接触限值 第1部分：化学有害因素》",
            "《危险化学品安全管理条例》（国务院令第591号）",
            "联网检索：各成分 GHS 分类、职业接触限值与毒理数据（见 rules/hazards_lib.md）",
        ],
    }

    os.makedirs(os.path.dirname(os.path.abspath(args.out)) or ".", exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(content, f, ensure_ascii=False, indent=2)
    def first_text(c):
        return c[0].get("text", "") if isinstance(c[0], dict) else c[0]
    n_pending = sum(1 for s in sections if s["content"] and first_text(s["content"]).startswith("【待"))
    print("[OK] content.json 初稿已生成（16节完整）:", args.out)
    print("  待推导节数: %d；自动识别成分: %d 项" % (n_pending, len(comps)))

if __name__ == "__main__":
    main()

