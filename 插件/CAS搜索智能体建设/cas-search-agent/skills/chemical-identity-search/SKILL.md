---
name: chemical-identity-search
description: 根据中文、英文、俗名、CAS 或 EC 号识别化学物质，并优先从官方/监管机构来源核验 CAS、EC、名称和法规信息；适用于水性材料、原料、助剂、单体、树脂、聚合物和法规筛查。
metadata:
  short-description: 中文物质名称的 CAS/EC 官方来源检索
---

# CAS/EC 官方来源搜索

将用户输入转换为可追溯的物质身份结论。主流程是 Agent 的多来源并行网页检索，脚本只负责生成检索计划、标准化候选和合并证据；PubChem 是辅助来源，不是唯一入口或主键。

## 工作规则

1. 先判断输入类型：中文名、英文名、俗名/商品名、CAS、EC/EINECS/ELINCS/List number、分子式、SMILES、InChI、InChIKey、PubChem CID 或混合物/聚合物描述。保留原始输入，不擅自把混合物、UVCB、盐类、异构体组合改写成单一物质。
2. 默认使用快速模式，并行检索至少五类来源：ECHA/ECHA CHEM、EPA CompTox、OECD eChemPortal、ChEBI/EMBL-EBI、NLM PubChem；输入涉及法规、REACH、职业健康、混合物或聚合物时切换深度模式，追加 EUR-Lex、NITE-CHRIP、GESTIS、CAS Common Chemistry、ChemSpider 和供应商 SDS。不得先等待 PubChem 失败后才开始其他来源。
3. 来源优先级：
   - A级：ECHA/ECHA CHEM、EPA/CompTox、其他政府或监管机构的化学品数据库、法规原文或官方清单。
   - B级：OECD eChemPortal、PubChem/NLM 等公共权威数据库，用于结构、同义名、CID 和交叉检索；它们不能单独证明 CAS Registry 权威性。
   - C级：供应商 SDS、行业数据库、百科或搜索摘要，只能作为线索，不能单独作为最终确认。
4. 对 CAS 和 EC 分别核验，不因为名称相同就默认两个编号都正确。优先引用能在页面正文或结构化字段中直接看到该编号的页面。
5. 如果官方页面与 PubChem、供应商 SDS 或本地数据库冲突，逐项列出冲突，不投票式强行选择；说明可能原因，如物质范围、盐型、异构体、群组条目、List number 或版本更新时间不同。
6. 对聚合物、树脂、UVCB、天然复杂物质、混合物、群组物质以及“及其盐类/异构体/组合”，优先返回范围说明和多个候选标识，并标注“不能可靠映射为单一 CAS/EC”。
7. 先使用本技能脚本生成保守变体和广域检索计划：NFKC 全半角归一化、去空格、括号/斜杠/顿号变体、CAS/EC 分隔符变体、中文原名、英文名、俗名、结构标识和 `site:` 查询式。然后让 Agent 并行执行多个来源查询；脚本中的 PubChem/Wikidata/ChEBI 适配器只是其中一部分，不得替代网页检索。
8. 不要把“查到更多”理解为接受模糊名称的所有结果。候选必须按 CID、InChIKey、结构、名称和来源路径去重；名称只有部分匹配、群组物质或结构不一致时，保留为低置信度候选并说明原因。
9. 每个来源必须记录状态：`verified`、`candidate`、`conflict`、`not-found` 或 `unavailable`。某个来源超时不代表物质不存在，也不能阻断其他来源。
10. 没有找到官方核验页时，必须明确写“官方来源未找到/当前无法访问”，并把 PubChem/Wikidata/ChEBI/SDS 等结果标为候选；不得把候选 CAS 写成已确认 CAS。

## 推荐工具

- 本技能目录下的 `scripts/chemical_lookup.py`：生成多来源搜索计划、输入变体、候选标准化和证据合并骨架；它不是唯一查询通道。
- 浏览器/网页搜索：负责并行查找实时官方页面；应保留最终页面 URL，不要只保留搜索结果页。优先使用官方域名限定的搜索式，并在同一轮发起多个来源查询。
- 详细来源路由见 [references/source-registry.md](references/source-registry.md)。
- 现有本地 REACH 数据库：如果用户要求法规筛查，可作为本地事实源，但必须同时给出数据库版本/路径和官方法规来源，不能把本地库替代官方页面。

## 输出格式

先给结论表，再给核验说明：

| 字段 | 结果 | 证据等级 | 来源/直接地址 | 状态 |
|---|---|---|---|---|
| 中文名称 |  | A/B/C |  | 已核验/候选/冲突 |
| 英文名称 |  | A/B/C |  |  |
| CAS |  | A/B/C |  |  |
| EC/List number |  | A/B/C |  |  |

随后列出：匹配依据、结构或同义名依据、来源覆盖矩阵、来源之间的差异、法规信息（仅在有相应官方来源时）、查询时间、仍需人工确认的事项。每个关键结论至少附一个可打开的直接来源地址；如果只有检索入口，必须标为“待打开核验”。

## 不可省略的免责声明

PubChem 官方说明其 CAS 覆盖不完整，且 PubChem 不对第三方提交的 CAS 进行人工权威核验。因此 PubChem CAS 只能作为候选或交叉证据，不能表述为 CAS Registry 的最终权威证明。ECHA InfoCard 也可能是群组或非法律约束性的摘要页面；引用时要说明页面性质和物质范围。
