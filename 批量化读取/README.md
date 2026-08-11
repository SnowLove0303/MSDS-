# MSDS 批量化读取

批量读取指定 MSDS Word 文件的入口与流程，复用 `结构读取/core` 的成熟 16 节解析
(GB/T 16483)，为数据清理 / 核对 / 入库提供统一批处理 CLI 与 Codex 技能。

## 目录结构

```
批量化读取/
├── batch_read.py                    # 核心入口 (Python CLI, 独立运行)
├── batch-read.ps1                   # PowerShell 封装 (自动定位 Python, 推荐)
├── README.md                        # 本文档
└── skills/
    └── msds-batch-read/
        └── SKILL.md                 # 技能文件 (给 Codex / Claude Code 使用)
```

## 快速开始

### PowerShell 封装 (推荐)

```powershell
.\batch-read.ps1 <docx或目录或通配符...> [选项]
```

### 直接 Python

```bash
python batch_read.py <docx或目录或通配符...> [选项]
```

脚本内部用绝对路径定位 `core/` 模块，不依赖当前工作目录。

## 典型用法

| 场景 | 命令 |
|---|---|
| 单文件提取 S9 | `.\batch-read.ps1 "F:\...\BL-8085 msds_CN 国彩.docx" --sections 9` |
| 目录递归提 S3 → TSV | `.\batch-read.ps1 "F:\数据库\MSDS\中文" --sections 3 --tsv --out s3.tsv --summary` |
| 通配符 + 检索 | `.\batch-read.ps1 "F:\数据库\MSDS\英文\*.docx" --query "供应商" --scope label` |
| **检索命中清单 (谁含某词)** | `.\batch-read.ps1 "F:\数据库\MSDS" --query "二丙二醇" --hits --summary` |
| **按型号过滤 + 检索** | `.\batch-read.ps1 "F:\数据库\MSDS" --query "供应商" --scope label --name-filter BL,OS --tsv` |
| **多关键词 OR** | `.\batch-read.ps1 "F:\数据库\MSDS" --query "危险 警示" --any --hits` |
| 全库报告 + 汇总 | `.\batch-read.ps1 "F:\数据库\MSDS" --report scan.json --summary --quiet` |
| 完整性校验 | `.\batch-read.ps1 "F:\数据库\MSDS" --verify --report integrity.json` |
| **文档×字段对比矩阵 → Excel** | `.\batch-read.ps1 "F:\数据库\MSDS\中文" --matrix --out 对比表.xlsx` |
| **矩阵 + 值状态三态列** | `.\batch-read.ps1 "F:\数据库\MSDS\中文" --matrix --states --out 对比表_带状态.xlsx` |
| **S3 成分矩阵 → TSV** | `.\batch-read.ps1 "F:\数据库\MSDS\中文" --sections 3 --matrix --out 成分矩阵.tsv` |
| **成分分列 → TSV** | `.\batch-read.ps1 "F:\数据库\MSDS\中文" --sections 3 --comp-cols --tsv --out 成分分列.tsv` |
| **成分分列 → JSON** | `.\batch-read.ps1 "F:\数据库\MSDS" --comp-cols --json --out 成分分列.json` |
| 环境自检 | `.\batch-read.ps1 doctor` |

完整参数表见 `batch_read.py` 头部文档注释，或 `.\batch-read.ps1 --help`。

## 批量化检索能力

跨文档关键词检索 (`--query`) 是批量化读取的核心能力，配合以下开关实现
"全库定位 / 按批次核对"：

| 开关 | 作用 |
|---|---|
| `--query 词` | 关键词检索 (全部文件统一; 多词空格分隔) |
| `--scope label\|value\|all\|section` | 检索范围; `section` = 精确节号 (如 `--query 9 --scope section`) |
| `--any` | 多关键词 **OR** (默认 AND: 全词命中) |
| `--name-filter 子串1,子串2` | 按文件名子串过滤 (逗号分隔或连续多值, OR, 大小写不敏感); 实现"只查某批次型号如 BL,OS" |
| `--hits` | **文件命中清单**: 每命中文件一行 `命中条目数<TAB>文件名`, 0 命中不列出, 末尾合计命中文件数/命中条目数 |
| `--show-empty` | 检索时保留 0 命中文件空块 (默认过滤, 减噪音) |
| `--summary` | 汇总表含 `🔍 检索 '词' (AND/OR): 命中文件 X/Y | 命中条目 Z` |

典型审计场景：`.\batch-read.ps1 "F:\数据库\MSDS" --query "二丙二醇" --hits --summary`
→ 直接列出哪些文件含该成分、各命中几处，0 命中文件不刷屏。

## 输出模型

三级父子级树（与 GUI 三列表格、覆写指向共用同一模型）：

```
节 (S1-S16 + S0 页眉页脚)
 └─ 大标题 (序号子标题, 如 "8.1 暴露控制" / "9.1 外观")
     └─ 字段 (无序号标签 + 内容; note 通栏; component 成分行)
```

- **text**: 每文件三级树 + 异常标注 + verify 结果
- **TSV**: `文件名 | 节 | 大标题 | 小标题 | 标签 | 内容` (带 BOM, Excel 可开)
- **JSON**: `{文件名: [条目]}` (stdout 或 --out)
- **矩阵** (`--matrix`): 行 = 一个文档, 列 = 字段; **列顺序 = (节号, 节内 reader
  上下顺序)**, 与 reader 三列表呈现一致。列头 = `Section{n} {序号范围} {标题}`
  (同标签不同序号自动合并且显示范围如 `9.4~9.5`; 无标题用 `Section{n} 特殊字段`;
  成分拆 名称/CAS/含量)。`.xlsx` 直接写 Excel (列头加粗/冻结首行/筛选/换行),
  `.tsv` 带 BOM 且换行折 ⏎
- **值状态三态列** (`--matrix --states`): 每字段列后加 `·状态` 列
  (有值=浅绿 / 无数据=浅灰 / 无此字段=浅红), 取代旧式 `-` 混淆
- **xlsx 美化**: 节分组表头 (值列深蓝 / 状态列紫色) + 三态着色 + 斑马纹 +
  细边框 + 列宽按内容自适应 + 冻结首行 + 自动筛选
- **成分分列** (`--comp-cols`): S3 成分从"每条成分一行、整格拼接"改为
  "每文档一行, 成分1|CAS1|含量1|成分2|CAS2|含量2 交替平铺" (text/TSV/JSON
  均适用; JSON 为 `{文件名: [{name, cas, conc}]}`), 方便数据库入库;
  只动输出层, GUI 显示与核心解析不变
- **报告**: 每文件状态 / summary / 各节统计 / 异常 / 缺失节 / (可选)条目 / verify

## 返回码

| 码 | 含义 |
|---|---|
| 0 | 全部成功 |
| 1 | 存在读取失败 或 输入无法匹配 (供批处理脚本判断) |
| 2 | 参数错误 / 未找到任何文件 |

## 给 Codex / Claude 安装技能

`skills/msds-batch-read/SKILL.md` 是通用技能文件（YAML frontmatter + Markdown）。

### 给 Codex

把技能目录复制到 Codex 技能目录：

```powershell
# 用户级 (所有项目可用)
Copy-Item -Recurse "F:\正式项目与模块化内容\Word 覆写模块\结构读取\批量化读取\skills\msds-batch-read" "$env:USERPROFILE\.codex\skills\"

# 或项目级
Copy-Item -Recurse "...\skills\msds-batch-read" "<你的项目>\.codex\skills\"
```

之后 Codex 会在用户要求批量处理 MSDS 时自动调用该技能
（description 声明了触发条件与入口路径）。

### 给 Claude Code

```powershell
Copy-Item -Recurse "...\skills\msds-batch-read" "$env:USERPROFILE\.claude\skills\"
```

## 标准数据库构建 (build_db.py)

批量读取的延伸: 把全库 MSDS 收敛为 SQLite 三表标准数据库, 供推断/分析直接查询。

| 表 | 说明 |
|---|---|
| `msds_documents` | 主表, 一行一文档 (S0/S1/S3.1/S4-S8/S10/S12-S14 稳定字段固定列) |
| `msds_components` | 成分表 1:N (成分数不固定, 用关联表正确建模) |
| `msds_fields` | 字段表 EAV (吸收 S9/S11 归一化字段, 160 种标签 → 43+ 标准字段) |

```bash
python build_db.py "F:\...\MSDS" -o MSDS数据库.sqlite   # 全库
python build_db.py "F:\...\MSDS" -o t.db --limit 20      # 前 20 个
```

设计文档见 `数据库/测试库/MSDS数据库设计.md`; 归一化映射表 `_NORM` 可审计可扩展。

## 全量信息导出到 Excel (export_excel.py)

把全库 646 份 MSDS 的**全部信息**汇总为一张宽表 Excel:

```bash
python export_excel.py [MSDS数据库.sqlite] [-o 输出.xlsx]
```

- **首行** = 字段标题, **首列** = 产品型号, 从**第二行**起写数据
- 行按**产品大类**(型号前缀, 如 OS/PU/PA/BL/BEK...) 从上到下分组排序, 类内按型号自然序
- **信息完整性保证**: S1-S16 全节字段
  - 标准字段(归一化) 单独成列, 列头 `S{n}·字段名`
  - 高频保留字段(覆盖≥15文档) 单独成列
  - 其余字段(含无标签正文) 并入该节 `S{n}·完整信息` 列 —— 零丢失
  - 成分 1..6 名称/CAS/含量 分列
  - `S3·成分原文` 列保留未结构化的文本成分完整原文 (英文模板 Chemicals 表格)
- **S2 信号词统一**: 9 种写法 (信号词/Signal word/警告词/警示词/警示词 危险/
  Warning word/Warning/Danger) → 单列, 160 文档覆盖; 误吞值自动纠正
- **S3 文本成分提取**: 英文模板 note 主要成分自动入成分列
- 实测 646 行 × 286 列, 字段表→Excel 0 缺失, 成分 5362 非空 0 缺失
- Excel 美化: 冻结首行首列 / 自动筛选 / 大类分组浅蓝底纹 / 自动换行 / 细边框

## 与主程序的关系

| 场景 | 用哪个 |
|---|---|
| 单文件可视化浏览 / 覆写指向分析 | `结构读取\main.py` (GUI) |
| 单文件命令行 16 节 | `.\Msds-Editor.ps1 cli <docx>` |
| 单/少量文件分层提取 | `.\Msds-Editor.ps1 extract <docx...>` |
| **批量读取 / 扫描 / 审计 (本目录)** | `.\batch-read.ps1 <输入...>` |

`batch_read.py` 与 `Msds-Editor.ps1 extract` 共享同一解析核心 (`core/`)，
批量化只是加上目录扫描 / 汇总统计 / 报告 / 完整性校验 / 统一返回码。