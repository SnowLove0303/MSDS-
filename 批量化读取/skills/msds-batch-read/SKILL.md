---
name: msds-batch-read
description: >
  批量读取 MSDS (化学品安全技术说明书) Word 文档并提取结构化内容。
  当用户需要一次处理多个 .docx 文件、递归扫描目录、按节(S1-S16)提取、
  按关键词跨文档检索(含文件命中清单/按型号过滤/多词AND·OR)、导出
  TSV/JSON、生成读取报告(含异常/缺失节/token完整性校验)时使用本技能。
  适用于英德市国彩精细化工 MSDS 数据清理/核对/入库场景。
  单文件阅读使用主程序 (main.py / Msds-Editor.ps1); 批量场景统一走 batch_read.py。
---

# MSDS 批量读取 (msds-batch-read)

## 能力

对一批 MSDS Word 文件 (.docx) 执行成熟的 16 节结构化解析 (GB/T 16483)，
输出统一三列表格模型 (节 → 大标题 → 字段+内容)，支持：

- **输入**: 单个文件 / 多个文件 / 目录 (递归扫描) / 通配符
- **节过滤**: 只提取指定节 (如 S3 成分表 / S9 理化性质)
- **关键词检索**: 按 label / value / 全字段 / 节号精确匹配过滤
- **多词组合**: 空格分隔多词, 默认 AND (全词命中), `--any` 切 OR (任一词)
- **按文件名/型号过滤** (`--name-filter`): 逗号分隔或连续多值, 大小写不敏感,
  实现"只检索某批次型号" (如 BL,OS 只查这两前缀型号)
- **文件命中清单** (`--hits`): 每命中文件一行 "命中条目数+文件名", 0 命中
  文件不列出, 末尾合计命中文件数/命中条目数; 批量审计"谁含某关键词"一键定位
- **检索命中统计**: `--summary` 汇总表含 `🔍 检索 '词' (AND/OR): 命中文件 X/Y
  | 命中条目 Z`; 检索时默认过滤 0 命中文件空块 (减噪音, `--show-empty` 保留)
- **输出格式**: 文本三级树 / TSV (Excel 可开) / JSON
- **宽表矩阵** (`--matrix`): 行 = 一个文档, 列 = 字段, 列顺序 = (节号, 节内 reader
  上下顺序); 列头 = "节 + 序号范围 + 标题" (如 `Section9 9.4~9.5 初沸点`, 同标签
  不同序号自动合并且显示范围); 无标题有文本用 `Section{n} 特殊字段`;
  成分拆 名称/CAS/含量 三列; `.xlsx` 直接写 Excel (列头加粗/冻结首行/筛选)
- **值状态三态列** (`--matrix --states`): 每字段列后加 `·状态` 列, 取值
  `有值` / `无数据`(原文有该字段但留空) / `无此字段`(原文没有), 取代旧式 `-` 混淆
- **成分分列** (`--comp-cols`): S3 成分从"每条成分一行、整格拼接"改为
  "每文档一行, 成分1\|CAS1\|含量1\|成分2\|CAS2\|含量2 交替平铺" (text/TSV/JSON
  均适用); 方便数据库入库. 只动输出层, GUI 显示与核心解析不变
- **数据库构建** (`build_db.py`): 全库 → SQLite 三表 (documents 主表 / components
  成分 1:N / fields 字段 EAV). **S1-S16 全节字段进字段表** (不再只 S9/S11),
  `_NORM` 归一化映射覆盖全节中英文/异写收敛 (S1 供应商/电话传真、S4 急救、
  S8 手部防护/暴露限值、S14 运输、S11 毒理等), 主表固定列中英别名匹配
  (英文模板全覆盖). 为推断/分析提供整齐标准数据库
- **全量导出 Excel** (`export_excel.py`): 全库 646 份 MSDS 全部信息 → 一张宽表
  (646 行 × 286 列). **首行字段标题, 首列产品型号, 第二行起写数据**, 行按
  产品大类(型号前缀 OS/PU/PA/BL...) 从上到下排序; 标准字段分列 + 低频字段并入
  `S{n}·完整信息` 列, **信息零丢失** (字段表→Excel 0 缺失); 成分 1..6
  名称/CAS/含量 分列; `S3·成分原文` 列保留未结构化的文本成分完整原文
  (英文模板 Chemicals 表格); Excel 冻结首行首列/筛选/大类分组底纹
- **S2 信号词统一**: 9 种写法 (信号词/Signal word/警告词/警示词/警示词 危险/
  Warning word/Warning/Danger) 归一化到 `信号词` 单列 (160 文档);
  Warning/Danger 标签被误吞的下一行值自动纠正
- **S3 文本成分提取**: 英文模板无标签 note 主要成分 (OS 系列 `· xxx`)
  自动入成分列 (名称/CAS/含量, 含量含 % / w/w 可拆分)
- **读取报告**: 每文件状态、字段/成分统计、异常清单、缺失节、token 完整性校验
- **返回码**: 0 全部成功 | 1 有失败或输入无法匹配 | 2 参数错误

解析器已兼容: 中英文模板、段落式文本、单列/多列成分表、自动编号序号恢复、
加粗标题归类、CAS 占位符 (商业机密 / Trade secret) 等 (见 结构读取/README.md)。

## 运行环境

- Python: `E:\MorenAnzhuangLujing\Anaconda\python.exe` (需 python-docx / openpyxl)
- 批量读取: `F:\正式项目与模块化内容\Word 覆写模块\结构读取\批量化读取\batch_read.py`
- 数据库构建: `F:\正式项目与模块化内容\Word 覆写模块\结构读取\批量化读取\build_db.py`
- 全量导出 Excel: `F:\正式项目与模块化内容\Word 覆写模块\结构读取\批量化读取\export_excel.py`
- 封装: `F:\正式项目与模块化内容\Word 覆写模块\结构读取\批量化读取\batch-read.ps1`
- 核心依赖: `F:\正式项目与模块化内容\Word 覆写模块\结构读取\core\` (docx_reader / extract / structure)

建议用 PowerShell 封装 (自动定位 Python, 无需手填解释器路径);
也可以直接 `python batch_read.py` (脚本内部已用绝对路径定位 core 模块)。

## 快速上手

```powershell
# 1. 单文件提取 S9 (理化性质)
.\batch-read.ps1 "F:\...\BL-8085 msds_CN 国彩.docx" --sections 9

# 2. 目录递归提取全部 S3 成分表 → TSV
.\batch-read.ps1 "F:\数据库\MSDS\中文" --sections 3 --tsv --out s3.tsv --summary

# 3. 通配符 + 关键词检索 (标签含"供应商")
.\batch-read.ps1 "F:\数据库\MSDS\英文\*.docx" --query "供应商" --scope label

# 4. 全库扫描 → JSON 报告 + 汇总 (不输出逐文件正文)
.\batch-read.ps1 "F:\数据库\MSDS" --report scan.json --summary --quiet

# 5. 全库 token 完整性校验 (慢, 逐文件与原文比对)
.\batch-read.ps1 "F:\数据库\MSDS" --verify --report integrity.json

# 6. 批量生成"文档 × 字段"对比矩阵 (宽表, 列头=Section{n} 序号 标题) → Excel
.\batch-read.ps1 "F:\数据库\MSDS\中文" --matrix --out 对比表.xlsx

# 7. 仅 S3 成分矩阵 (成分名/CAS/含量拆三列) → TSV
.\batch-read.ps1 "F:\数据库\MSDS\中文" --sections 3 --matrix --out 成分矩阵.tsv

# 8. 矩阵 + 值状态三态列 (有值/无数据/无此字段) → Excel
.\batch-read.ps1 "F:\数据库\MSDS\中文" --matrix --states --out 对比表_带状态.xlsx

# 9. 成分分列输出 (每文档一行, 成分1|CAS1|含量1|成分2|... 交替平铺) → TSV
.\batch-read.ps1 "F:\数据库\MSDS\中文" --sections 3 --comp-cols --tsv --out 成分分列.tsv

# 10. 成分分列 JSON (数据库入库用: {文件名: [{name, cas, conc}]})
.\batch-read.ps1 "F:\数据库\MSDS" --comp-cols --json --out 成分分列.json

# 11. 全库全部信息汇总 → Excel 宽表 (首行字段标题, 首列产品型号, 按大类排序)
python "F:\...\批量化读取\export_excel.py" "F:\...\MSDS数据库.sqlite" -o "MSDS全量汇总.xlsx"

# ---- 批量化检索增强 ----
# 12. 文件命中清单: 哪些文件含"二丙二醇", 各命中几处 (批量审计"谁命中")
.\batch-read.ps1 "F:\数据库\MSDS" --query "二丙二醇" --hits --summary

# 13. 只查某批次型号 + 检索 (BL/OS 型号里查"供应商"标签) → TSV
.\batch-read.ps1 "F:\数据库\MSDS" --query "供应商" --scope label --name-filter BL,OS --tsv

# 14. 多关键词 OR (命中任一词即算; 默认 AND 全词命中)
.\batch-read.ps1 "F:\数据库\MSDS" --query "危险 警示" --any --hits

# 15. 按节号精确检索 (scope=section 精确匹配, 等价于按节筛)
.\batch-read.ps1 "F:\数据库\MSDS\中文" --query 9 --scope section --tsv --out s9.tsv
```

## 参数

| 参数 | 说明 |
|---|---|
| `<docx或目录或通配符...>` | 位置参数, 可多个; 目录自动递归, 自动跳过 Word 锁文件 `~$*` |
| `--sections 1,3,9` | 仅提取指定节 (逗号分隔) |
| `--query 词` | 关键词检索 (全部文件统一; 多词空格分隔, 默认 AND) |
| `--scope label\|value\|all\|section` | 检索范围, 默认 all; section=精确节号 (如 `--query 9 --scope section`) |
| `--any` | 多关键词 OR 匹配 (默认 AND: 全词命中) |
| `--name-filter 子串1,子串2` | 按文件名子串过滤待处理文件 (逗号分隔或连续多值, OR, 大小写不敏感); 实现"只查某批次型号如 BL,OS" |
| `--hits` | 文件命中清单: 每命中文件一行 `命中条目数\t文件名`, 0 命中不列出, 末尾合计; 配合 `--query` 做批量审计 |
| `--show-empty` | 检索时保留 0 命中文件空块 (默认过滤, 减噪音) |
| `--json` / `--tsv` / `--matrix` | 输出格式 (默认 text 三级树); TSV 带 BOM 可直接 Excel 打开 |
| `--out 文件` | 导出到文件 (不受 --quiet 抑制); `--matrix` 且 `.xlsx` 时写 Excel 工作簿 |
| `--states` | 仅 `--matrix`: 每字段列后加"值状态"三态列 (有值/无数据/无此字段) |
| `--comp-cols` | 成分分列输出: 每文档一行, 成分1\|CAS1\|含量1\|成分2\|... 交替平铺 (text/TSV/JSON) |
| `--summary` | 打印汇总统计表 (成功/失败/字段/成分/异常分布/缺失节) |
| `--report 文件` | 生成 JSON 报告 (每文件状态/异常/缺失节/节覆盖) |
| `--with-entries` | 报告含提取条目全文 (默认剥离, 控文件体积) |
| `--verify` | token 级完整性校验 (慢) |
| `--fail-fast` | 首个读取失败立即退出 |
| `--skip-empty` | 跳过无内容文件 |
| `--verbose` | 打印每文件处理进度 |
| `--quiet` | 只输出汇总/报告 |

## 常见任务

### 任务 1: 批量核对某批 MSDS 的 S3 成分表
```powershell
.\batch-read.ps1 "F:\数据库\MSDS\中文" --sections 3 --tsv --out 成分表.tsv --summary
```
TSV 列: 文件名 | 节 | 大标题 | 小标题 | 标签 | 内容。成分行内容形如
`成分名 | CAS: xxx | 含量: y%`, 占位符 CAS (商业机密/Trade secret) 原样保留。

### 任务 2: 找出缺失节 / 有异常的 MSDS
```powershell
.\batch-read.ps1 "F:\数据库\MSDS" --report scan.json --summary --quiet
```
汇总表列出缺失节文件与异常分布; 报告 JSON 里 `files[].missing_sections` /
`files[].anomalies` 可程序化消费。

### 任务 3: 批量提取某字段值 (如 S1 产品名称)
```powershell
.\batch-read.ps1 "F:\数据库\MSDS\英文" --sections 1 --query "Product name" --scope label --tsv --out names.tsv
```

### 任务 4: 完整性审计 (token 级, 确认无内容遗漏)
```powershell
.\batch-read.ps1 "F:\数据库\MSDS" --verify --report integrity.json --summary --quiet
```
报告里 `files[].verify.missing_count` = 该文件解析结果与原文逐 token 比对的遗漏数,
>0 表示有内容未解析出来 (需回查解析器兼容性)。

### 任务 5: 生成"文档 × 字段"对比矩阵 (与 reader 显示同构)
```powershell
.\batch-read.ps1 "F:\数据库\MSDS\中文" --matrix --out 对比表.xlsx --summary
```
- 每行 = 一个文档, 每列 = 一个字段; **列顺序 = (节号, 节内 reader 上下顺序)**,
  与 reader 三列表呈现完全一致。
- 列头 = reader 三列表里的"节 + 序号范围 + 标题":
  - 有序号字段 → `Section9 9.4~9.5 初沸点` (同标签不同序号自动合并, 显示覆盖范围)
  - 无序号字段 → `Section2 GHS分类`
  - 无标题有文本 → `Section5 特殊字段`
  - 成分 → `Section3 成分1-名称/CAS/含量` (按索引拆三列)
- 同 (节, 标签) 不同序号合并为一列 → 即使不同模板字段序号顺延
  (如有的多一个 `9.4 离子性` 导致初沸点 9.4→9.5), 列也不会分裂, 只在列头显示范围。
- 输出 .xlsx: **已美化** — 节分组表头 (值列深蓝 / 状态列紫色白字加粗) +
  冻结首行 + 自动筛选 + 单元格自动换行 + 细边框 + 斑马纹行 +
  列宽按内容自适应 (中文按 2 宽估算);
  输出 .tsv: 带 BOM, 换行折成 ⏎ 保证一行一文档。

### 任务 6: 矩阵 + 值状态三态列 (区分"无数据"与"无此字段")
```powershell
.\batch-read.ps1 "F:\数据库\MSDS\中文" --matrix --states --out 对比表_带状态.xlsx
```
- 每个字段列后跟一个 `·状态` 列, 取值三态:
  - `有值` — 原文有内容 (浅绿底 + 绿字)
  - `无数据` — 原文有该字段但留空, 如 `1.3 供应商信息：` 空值 (浅灰底)
  - `无此字段` — 原文根本没有该字段, 如 BL-8085 无 `9.4 离子性` (浅红底 + 红字)
- xlsx 中三态已着色, 一眼可辨; 取代旧式 `-` 填满 (把"无此字段"和"无数据"混为一谈) 的做法。

### 任务 7: 成分分列输出 (数据库入库)
```powershell
# TSV: 每文档一行, 列 = 成分1|CAS1|含量1|成分2|CAS2|含量2|... (最多成分数定列数)
.\batch-read.ps1 "F:\数据库\MSDS\中文" --sections 3 --comp-cols --tsv --out 成分分列.tsv

# JSON: {文件名: [{name, cas, conc}, ...]} 直接可入库
.\batch-read.ps1 "F:\数据库\MSDS" --comp-cols --json --out 成分分列.json

# 配合 --query 检索: 命中成分的文件也按分列输出
.\batch-read.ps1 "F:\数据库\MSDS" --comp-cols --json --query "二丙二醇"
```
- 只动**输出层**; reader GUI 显示 (三列表格) 与核心解析 (core/docx_reader、
  core/structure 的 ComponentData) 完全不变。
- 默认输出不带 `--comp-cols` 时保持原样 (每条成分一行、整格拼接), 互不影响。

### 任务 8: 批量化检索 — 全库定位含某内容的文件
```powershell
# 1. 哪些文件含"二丙二醇", 各命中几处 (命中清单, 按命中数降序)
.\batch-read.ps1 "F:\数据库\MSDS" --query "二丙二醇" --hits --summary
#   输出: 命中条目数<TAB>文件名 ... 末尾 "命中文件 X / Y | 命中条目 Z"

# 2. 只查某批次型号 (BL/OS) 里标签含"供应商"的字段 → TSV 核对
.\batch-read.ps1 "F:\数据库\MSDS" --query "供应商" --scope label \
    --name-filter BL,OS --tsv --out bl_os_供应商.tsv

# 3. 多关键词 OR (成分检索: 含"二丙二醇"或"乙二醇丁醚"任一)
.\batch-read.ps1 "F:\数据库\MSDS" --query "二丙二醇 乙二醇丁醚" --any --hits

# 4. 检索+报告: 命中文件清单和 JSON 报告同时出 (报告默认不含条目全文)
.\batch-read.ps1 "F:\数据库\MSDS" --query "供应商" --scope label \
    --hits --summary --report scan_检索.json --quiet
```
- **命中含义**: 检索时条目数 = 该文件命中条数; 0 命中的文件默认不输出
  (text/hits 均过滤, `--show-empty` 保留), 避免噪音.
- **scope=section 精确节号**: `--query 9 --scope section` 只返回第 9 节字段;
  若只是想筛第 9 节, 直接用 `--sections 9` 更直观.

### 任务 9: 构建标准数据库 (为推断/分析)
```bash
# 全库 646 → SQLite 三表 (documents/components/fields)
python "F:\正式项目与模块化内容\Word 覆写模块\结构读取\批量化读取\build_db.py" \
  "F:\正式项目与模块化内容\Word 覆写模块\数据库\MSDS" -o MSDS数据库.sqlite

# 小批量试建
python build_db.py "...\MSDS" -o t.db --limit 20
```
- **主表 msds_documents**: 一行一文档, S0 页眉 + S1/S4-S8/S10/S12-S14 稳定字段固定列
  (中英别名匹配, 英文模板全覆盖)
- **成分表 msds_components**: 1:N, 成分数不固定 (实测 1~6 个) 用关联表正确建模
- **字段表 msds_fields**: EAV, **吸收 S1-S16 全节字段**; 字段名经 `_NORM`
  归一化 (中英文/异写收敛: 安全储存条件 ↔ Safe storage conditions 等)
- **归一化**: 全节中英收敛, 标准字段覆盖 76.6%; 未命中多为合理保留的值字段
  (毒理数据表列/手套材料选项/EU 生态数据项)

## 输出解读

- 汇总表 `异常分布 (按节)`: S3 的拆分/全角归一告警是**预期兼容性动作**, 非错误;
  S9 编号不连续 = 文档自动编号 + 显式编号混排 (缺失号是文档真实情况)。
- 返回码供批处理脚本 (for / powershell / CI) 判断成败: 0 成功, 1 有失败。

## 注意事项

1. **绝对路径**: 所有路径建议用绝对路径, 脚本不依赖当前工作目录。
2. **~$ 临时文件**: Word 打开中的文档会生成 `~$xxx.docx`, 已自动跳过。
3. **非 .docx**: 只处理 .docx; 老版 .doc 需先转 docx (可用 Word 或批量转换工具)。
4. **编码**: 输出默认 UTF-8; TSV 带 BOM 供 Excel; 报告 JSON 为 UTF-8。
5. **速度**: 646 文件全库扫描约 30-40s (实测 35s, 每文件 ~50ms, 含 docx 解压
   与 16 节解析); 只跑子集用 `--name-filter`/通配符收窄可显著加速; --verify
   约慢 2-3 倍。检索本身零负担 (单次全库检索 ~70ms)。
6. **报告体积**: 默认报告不含条目全文 (全库含数万条目 → 数十 MB), 需要时加 --with-entries。
7. **不要修改源文件**: batch_read.py 只读不写, 与核心解析器行为一致。