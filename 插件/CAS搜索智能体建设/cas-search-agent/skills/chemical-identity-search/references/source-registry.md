# 广域来源注册表

技能使用时应根据输入类型选择来源并行检索。来源域名只是路由提示，最终必须打开页面确认页面内容、物质范围和编号。

## A 级：官方/监管优先

| 来源 | 入口 | 重点 |
|---|---|---|
| ECHA Substance Information | https://echa.europa.eu/substance-information | 物质名、EC/CAS、REACH 信息 |
| ECHA CHEM | https://chem.echa.europa.eu/ | ECHA 最新监管数据和物质页面 |
| EPA CompTox | https://comptox.epa.gov/dashboard/ | EPA 维护的化学品名称、CASRN、DTXSID、结构 |
| OECD eChemPortal | https://www.echemportal.org/echemportal/substance-search | 多个政府/监管数据集的聚合检索 |
| EU EUR-Lex | https://eur-lex.europa.eu/ | 法规正文、官方清单和物质范围 |
| NITE-CHRIP | https://www.nite.go.jp/en/chem/chrip/chrip_search/systemTop | 日本官方化学品/法规信息 |
| GESTIS | https://gestis.dguv.de/ | 德国职业安全化学品信息和标识 |
| 中国政府/监管来源 | https://www.gov.cn/ | 仅在页面明确给出物质标识或法规条目时使用 |

## B 级：公共权威/专业数据库

| 来源 | 入口 | 重点 |
|---|---|---|
| ChEBI / EMBL-EBI | https://www.ebi.ac.uk/chebi/ | 标准名称、同义名、结构和交叉标识 |
| NLM PubChem | https://pubchem.ncbi.nlm.nih.gov/ | 结构、同义名、CID、交叉标识；CAS 仅作候选 |
| Wikidata | https://www.wikidata.org/ | 中文/英文标签、别名和交叉标识桥接 |
| ChemSpider | https://www.chemspider.com/ | 名称/结构/同义名补充，需交叉核验 |
| CAS Common Chemistry | https://commonchemistry.cas.org/ | 若页面可公开访问，可作为 CAS 直接核验优先来源 |

## C 级：线索来源

供应商 SDS、制造商技术资料、行业数据库、搜索摘要和商品页可用于发现英文名、俗名和可能 CAS，但不能单独确认 CAS/EC。SDS 必须记录制造商、版本日期、物质范围和原始 URL。

## 快速模式与深度模式

- 快速模式：并行查询 ECHA、EPA CompTox、PubChem、ChEBI、OECD eChemPortal 五类来源；通常先返回候选和可核验地址。
- 深度模式：在快速模式基础上加入 EUR-Lex、NITE-CHRIP、GESTIS、CAS Common Chemistry、ChemSpider、供应商 SDS 和本地法规库，并逐字段处理冲突。
- 任何来源超时都不能阻断其他来源；结果中记录 `unavailable`、`timeout` 或 `not-found`，而不是静默丢弃。
