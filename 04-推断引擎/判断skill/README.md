# 判断skill — MSDS 规范化推导生成

本项目内维护的 `msds-inference-write` 技能副本（数据驱动：规则库 + 通用脚本 + 模板）。

## 交付约定（每次使用必须满足）
1. **必须产出两份 Word**：每次调用结束必须交付 .docx 方案文档：
   - **方案 Word**（默认）：含封面/诊断汇总表/推导依据/16节（【原文问题】+【规范应写内容】）；
     仅说明内容（推导理由/编辑指令/字段提醒）以灰色【说明】标出。
   - **纯 MSDS 正文 Word**（`--pure`）：只抽取"可写入 MSDS"的内容，形成一份干净可用的 SDS 正文。
2. **交付位置**：`<源MSDS所在目录>\推导方案\<产品名> MSDS规范化推导方案.docx`
   及 `<产品名> MSDS正文.docx`
3. **质量对标黄金范例**：`templates/os1330-full-example.json` —— 16 节全部有内容、
   每节 =【原文问题】+【规范应写内容】、含封面/诊断表/依据/参考、content 条目区分 msds/note。
   典型交付样例：`入库word  第一批\推导方案\OS-1330 MSDS规范化推导方案.docx`

## 位置关系
- 本项目副本：本目录（`推断引擎/判断skill/`）
- 用户级副本（所有项目可用）：`C:\Users\52882\.claude\skills\msds-inference-write\`
- **打磨时建议以本项目副本为准**：改动后同步回用户级副本，或在此维护后由用户决定。

## 快速使用（在本项目内）
```bash
# 1. 读取源 MSDS → facts.json（自动识别16节，重点抽 S1/S3/S9）
python "推断引擎\判断skill\scripts\msds_reader.py" "<源MSDS.docx>" --out "推导方案\facts.json"

# 2. 生成 16 节完整初稿 → content.json（S3成分自动解析+速查库匹配，其余节占位待完善）
python "推断引擎\判断skill\scripts\build_initial_content.py" "推导方案\facts.json" \
       --product "<产品名>" --out "推导方案\content.json"

# 3. AI 联网检索成分危害数据，逐节完善 content.json（直到16节无"待推导"占位）

# 4. 生成规范 Word（每次产出两份）
python "推断引擎\判断skill\scripts\msds_docx.py" "推导方案\content.json" \
       --out "推导方案\<产品名> MSDS规范化推导方案.docx"
python "推断引擎\判断skill\scripts\msds_docx.py" "推导方案\content.json" --pure \
       --out "推导方案\<产品名> MSDS正文.docx"
```

## 打磨入口
- `rules/limits.md` —— 混合物危害组分浓度限值（推导判据核心）
- `rules/hazards_lib.md` —— 常用成分危害速查库（扩充为主）
- `rules/output_structure.md` —— Word 输出结构与每节必填字段
- `templates/os1330-full-example.json` —— **黄金范例**（16节完整推导方案，质量基准）
- `templates/content.example.json` —— 输出内容格式示例
- `SKILL.md` —— 技能入口、交付约定与 16 节铁律表

> 脚本（scripts/）一般不需改动；如需改动，先在本副本验证再同步回用户级副本。
