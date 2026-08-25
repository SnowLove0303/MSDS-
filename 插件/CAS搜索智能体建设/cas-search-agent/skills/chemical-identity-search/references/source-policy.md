# 官方来源与链接策略

## 来源入口

- ECHA Substance Information / InfoCard: https://echa.europa.eu/substance-information
- ECHA CHEM: https://chem.echa.europa.eu/
- EPA CompTox Chemicals Dashboard: https://comptox.epa.gov/dashboard/
- OECD eChemPortal: https://www.echemportal.org/echemportal/substance-search
- PubChem: https://pubchem.ncbi.nlm.nih.gov/
- PubChem PUG REST 文档: https://pubchem.ncbi.nlm.nih.gov/docs/pug-rest
- ChEBI/EMBL-EBI REST API: https://www.ebi.ac.uk/chebi/tools
- EMBL-EBI Ontology Lookup Service: https://www.ebi.ac.uk/ols4/

## 判定原则

官方域名不等于每个字段都经过同一机构确认。记录页面中实际出现的字段、页面性质、更新时间和物质范围。ECHA 的 EC 可能是 EC Inventory 编号，也可能是 List number；两者必须在输出中区分。法规清单中的群组物质可能没有单一 CAS，不能用一个方便的 CAS 替代整个群组。PubChem 的 identifier type 查询优先于从普通 synonym 文本中正则提取，因为 PubChem 官方说明前者保留了原始来源组织的标识关系，而普通 synonym 可能包含未经验证的第三方标识。

## URL 要求

优先保存具体物质页或具体法规条目页。只有在尚未找到具体页时才保存搜索入口，并明确标记为待核验。不要把搜索摘要、供应商商品页或未经核对的聚合页面作为唯一来源。
