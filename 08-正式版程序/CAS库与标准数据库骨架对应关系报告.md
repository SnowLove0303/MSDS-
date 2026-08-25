# 检索工作台 CAS 库与标准数据库骨架对应关系报告

## 1. 核验范围

- Web 入口：`http://127.0.0.1:8765/?v=639229961953086635`
- 项目根目录：`F:\正式项目与模块化内容\冠志\MSDS\08-正式版程序`
- 核验方式：读取项目源码、SQLite 元数据与当前运行服务的只读 API；未修改数据库、源码或服务状态。
- 核验时间：2026-08-22。

### 2026-08-22 CAS Section 2 补充更新

本次已按固定 13 项骨架为正式 CAS 主库中 50 个明确 CAS 建立/补充当前 profile，并写入正式 `cas_section2_result.db`：50 个 profile、650 条骨架字段记录、50 条 ChatGPT Web 辅助资料来源；550 条非身份字段保留为 `derived_candidate`，50 个 profile 全部保持 `manual_review`。59 条无 CAS 记录未生成 CAS profile。原正式结果库已备份至 `F:\正式项目与模块化内容\冠志\MSDS\_codex_backup\cas_s2_enrichment_20260822_145724\cas_section2_result.db`，可用于回滚。

## 2. 结论

CAS 库检索工作台使用的是正式物质库 `cas_library.db`。页面在 `assets/app.js` 的“检索工作台 → CAS 库检索”分支中调用 `/api/cas`，服务端 `server.py` 的 `_cas_search()` 直接读取该 SQLite 库，并返回物质、别名、型号使用关系以及对应的 Section 2 标准结果。

“CAS 库”和“CAS 标准数据库骨架”是两个相互关联但职责不同的数据库：

1. `cas_library.db`：CAS 身份、标准名称、别名和型号使用关系主库。
2. `cas_section2_result.db`：CAS 物质级 Section 2 标准结果、来源、条件、分类、标签、危害摘要和结构化事实库；其中 `cas_s2_skeleton_field` 是固定的 13 字段标准输出骨架。
3. `cas_field_mapping.db`：CAS/无 CAS 成分身份及原始名称字段映射库，不是 Section 2 结果骨架。

## 3. 文件与职责对应

| 对象 | 绝对路径 | 作用 | 当前规模 |
|---|---|---|---:|
| Web 前端 | `F:\正式项目与模块化内容\冠志\MSDS\08-正式版程序\assets\app.js` | CAS 检索 UI、请求 `/api/cas`、渲染详情和 Section 2 标准结果 | 52,895 bytes |
| Web 服务 | `F:\正式项目与模块化内容\冠志\MSDS\08-正式版程序\server.py` | 定义数据库路径、CAS 查询、Section 2 结果读取和 `/api/cas` 路由 | 40,382 bytes |
| 配置 | `F:\正式项目与模块化内容\冠志\MSDS\08-正式版程序\db_config.json` | 将 `cas_library.db` 显示为“CAS 库”，业务树标识为 `cas`；将映射库标识为 `cas_mapping` | 440 bytes |
| CAS 主库 | `F:\正式项目与模块化内容\冠志\MSDS\03-数据库\正式库\Data Base\cas_library.db` | CAS 物质主数据和型号关联 | 270,336 bytes；4 表 |
| CAS 字段映射库 | `F:\正式项目与模块化内容\冠志\MSDS\03-数据库\正式库\Data Base\cas_field_mapping.db` | CAS 身份、原始成分名、异写、冲突和观察关系 | 401,408 bytes；5 表 |
| Section 2 结果库 | `F:\正式项目与模块化内容\冠志\MSDS\03-数据库\正式库\Data Base\cas_section2_result.db` | CAS 物质级标准结果与证据数据 | 221,184 bytes；10 表及元数据/视图 |
| 标准骨架定义 | `F:\正式项目与模块化内容\冠志\MSDS\04-推断引擎\推断引擎程序\cas_section2_result_schema.sql` | 创建结果库结构、13 字段约束、固定骨架触发器和标准视图 | 16,646 bytes |

## 4. CAS 主库结构

数据库：`03-数据库\正式库\Data Base\cas_library.db`

| 表 | 行数 | 关键职责 |
|---|---:|---|
| `cas_substance` | 109 | `cas_id`、CAS 号、标准名称、类别、注册状态、来源数、型号数、备注 |
| `cas_alias` | 130 | 一个物质对应的原始/别名表述及出现次数 |
| `cas_model_usage` | 804 | CAS 与产品型号、原始名称、原始 CAS、浓度、来源文件的关联 |
| `cas_note` | 2 | 物质级备注、说明及来源 |

`cas_library.db` 的主身份表是 `cas_substance`；别名和型号关系通过 `cas_id` 外键关联。Web 查询支持 CAS 号、标准名称、类别、注册状态、备注、别名、型号名、原始名称和原始 CAS 的模糊检索，并支持 `has_cas`、`no_cas`、`secret` 状态筛选。

## 5. 字段映射库结构

数据库：`03-数据库\正式库\Data Base\cas_field_mapping.db`

| 表 | 行数 | 关键职责 |
|---|---:|---|
| `cas_identity` | 112 | CAS/无 CAS 成分身份基准及标准名称 |
| `cas_field_mapping` | 157 | 原始名称到标准名称/身份键的映射 |
| `cas_mapping_observation` | 804 | 映射在型号、原始 CAS、浓度和来源文件中的观察记录 |
| `cas_mapping_conflict` | 13 | 一个原始名称对应多个候选身份时的冲突记录 |
| `cas_mapping_meta` | 6 | 映射库元数据 |

该库由 Web 的“CAS 字段映射库”条目展示，服务端配置路径为 `CAS_MAPPING_DB_PATH`。它支撑成分名称标准化和冲突追踪，但当前 CAS 检索详情的主列表仍来自 `cas_library.db`。

## 6. Section 2 标准数据库骨架

数据库：`03-数据库\正式库\Data Base\cas_section2_result.db`

定义文件：`04-推断引擎\推断引擎程序\cas_section2_result_schema.sql`

核心数据表包括：

- `cas_s2_profile`：CAS 当前物质档案和版本身份。
- `cas_s2_source`：来源文件、来源哈希、来源级别、法规版本和证据定位。
- `cas_s2_condition`：含量阈值、适用范围和条件状态。
- `cas_s2_classification`：GHS/危险性分类结果。
- `cas_s2_label_element`：象形图、信号词、危险性说明和防范说明。
- `cas_s2_hazard_summary`：物理与化学、健康、环境及其他危害摘要。
- `cas_s2_fact`：结构化事实及其适用范围、数值和证据状态。
- `cas_s2_review`：人工复核状态和复核记录。
- `cas_s2_skeleton_field`：固定 13 字段标准骨架。
- `cas_s2_result_field`：每个 CAS profile 对应的 13 项标准输出值。

当前库的 `schema_meta` 标记如下：

| 元数据 | 值 |
|---|---|
| `schema_id` | `cas_section2_result_library` |
| `schema_version` | `1.2.0` |
| `standard_skeleton_id` | `nmp_sheet1_13_fields` |
| `standard_skeleton_field_count` | `13` |
| `standard_skeleton_missing_value` | `无数据` |
| `identity_scope` | CAS 身份库独立维护；本库只保存 CAS Section 2 结果事实 |

### 固定 13 字段

| 顺序 | 字段键 | 中文标签 | 类型 |
|---:|---|---|---|
| 1 | `standard_name` | 标准名称 | identity |
| 2 | `cas_no` | CAS | identity |
| 3 | `ec_no` | EC | identity |
| 4 | `index_no` | Index No | identity |
| 5 | `ghs_classification` | GHS 危险性类别 | classification |
| 6 | `pictogram` | 象形图 | label |
| 7 | `signal_word` | 信号词 | label |
| 8 | `hazard_statement` | 危险性说明 | label |
| 9 | `precautionary_statement` | 防范说明 | label |
| 10 | `physical_chemical_hazard` | 物理与化学危险 | summary |
| 11 | `health_hazard` | 健康危害 | summary |
| 12 | `environmental_hazard` | 环境危害 | summary |
| 13 | `other_hazard` | 其他危害 | summary |

SQL 骨架通过触发器限制 `cas_s2_skeleton_field` 只能保留这 13 个字段，禁止修改或删除；`v_cas_s2_current_standard_result` 通过“当前 profile × 13 字段”交叉连接，保证即使某项没有实际值，Web 输出仍能显示该标准字段，并以“无数据”补齐。

## 7. Web 调用链

```text
检索工作台
  └─ assets/app.js：选择“CAS 库检索”
      └─ GET /api/cas?q=...&status=...
          └─ server.py：_cas_search()
              ├─ cas_library.db / cas_substance
              ├─ cas_alias
              ├─ cas_model_usage
              └─ _cas_s2_result(cas_no)
                  └─ cas_section2_result.db
                      └─ v_cas_s2_current_standard_result
                          └─ standard_fields[13]
```

页面详情渲染逻辑在 `assets/app.js` 中将 `s2_result.standard_fields` 按 `display_order` 排序，并显示为“Section 2 标准结果”。若对应 CAS 没有当前 profile，页面显示“该 CAS 暂无标准骨架结果”；如果有 profile，则固定输出 13 项字段。

## 8. 运行态交叉核验

当前服务 `http://127.0.0.1:8765` 返回正常：

- `/api/health`：`ok=true`。
- `/api/libraries`：明确返回“CAS 库”和“CAS 字段映射库”的正式路径及统计。
- `/api/cas`：返回 109 条 CAS 物质。
- `/api/cas?q=872-50-4`：命中 NMP，并返回 13 项标准字段；当前结果库中的 profile 状态为 `manual_review`。
- `/api/cas?q=2687-91-4`：命中 N-乙基吡咯烷酮，并返回 13 项标准字段；当前结果库中的 profile 状态为 `manual_review`。

## 9. 注意事项

1. `cas_library.db` 与 `cas_section2_result.db` 必须按 CAS 号关联；前者负责物质/型号检索，后者负责 Section 2 标准结果。
2. `cas_field_mapping.db` 不应被当作 CAS 主库或 Section 2 结果库。
3. 当前正式 Web 服务使用 `server.py` 的旧式 HTTP 路由 `/api/cas`；项目中的 `frontend/src/api.js` 还存在 FastAPI 风格的 `/api/libraries/cas/substances` 路由，但当前 8765 服务实测该路径返回 404。因此，当前页面实际生效入口是 `assets/app.js → /api/cas`。
4. `cas_library.db` 在计算 SHA-256 时被正在运行的服务占用，未强行复制、停止服务或改变文件状态；本报告保留文件大小、修改时间和数据库结构作为核验依据。

## 10. 结论性定位

若要定位“CAS 库”：使用

`F:\正式项目与模块化内容\冠志\MSDS\03-数据库\正式库\Data Base\cas_library.db`

若要定位“CAS 标准数据库骨架”：使用

`F:\正式项目与模块化内容\冠志\MSDS\03-数据库\正式库\Data Base\cas_section2_result.db`

并以

`F:\正式项目与模块化内容\冠志\MSDS\04-推断引擎\推断引擎程序\cas_section2_result_schema.sql`

作为结构定义和固定 13 字段骨架的权威来源。

## 11. Section 2 骨架换行规范化

2026-08-22 已对 50 个已有 CAS 的 Section 2 结果执行隔离暂存、备份和正式库回写。针对多项内容，将 GHS 分类、象形图、信号词、危险性说明、防范说明及物理/健康/环境/其他危害字段统一保存为“一条一行”的实际换行格式；CAS、标准名称、EC、Index No 等身份字段未修改。

- 规范化字段：192 个。
- 当前 profile：50 个。
- 当前结果字段：650 个。
- 实际包含换行的字段：393 个。
- 字面量 `\\n`：0 个。
- 目标字段中的分号分隔多项值：0 个。
- 正式库与暂存库 SHA-256：`fc1abf888ea559574481d84ccefa5050f4cf67dfa41aa5e8f0b0b90dc22cb9f2`。
- 回滚备份：`F:\正式项目与模块化内容\冠志\MSDS\_codex_backup\cas_linebreak_fix_20260822_165120\cas_section2_result.db`。

前端 `assets/app.js` 继续按固定 13 项字段渲染，`assets/styles.css` 的 `.cas-result-table` 保留 `white-space: pre-wrap`，因此数据库中的实际换行会在检索工作台中按规范显示。

## 12. GHS 象形图字典与 CAS 检索展示

2026-08-22 已将 9 个标准 GHS 象形图写入 `cas_section2_result.db` 的 `cas_s2_pictogram_catalog` 表，保存编号、中文含义、PNG 文件名、MIME 类型、图片二进制、SHA-256 和来源路径。当前图片二进制总量为 331,697 字节。

服务端 `/api/ghs/pictograms` 已优先从该数据库表读取图片并返回同源 Data URI；CAS 检索详情页已按“象形图”及兼容名称“GHS象形图”渲染编号、中文含义和图片。条件性文本仍保留在原字段中，不会将条件性说明误当作无条件图标。
