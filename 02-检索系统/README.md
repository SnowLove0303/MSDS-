# MSDS 结构读取 · 检索系统

以固定冠志中文模板为标准模板，提供 **导入 → 读取 → 显示 → 检索 → 表单覆写 → Word 产出** 流程。
按 GB/T 16483 把 MSDS Word 文件解析为 S0 页眉/页脚 + S1~S16 的 17 节父子级结构（节 → 大标题/小标题 → 字段），支持全库/单文件关键词检索与批量提取。

## 运行

### 直接运行

```bash
python main.py                     # 启动可视化界面 (左目录 | 右表格 + 检索框)
python main.py --cli 文件.docx     # 命令行解析并打印 16 节
python main.py --extract 文件.docx # 分层检索提取
```

### PowerShell 封装 (Msds-Editor)

```powershell
.\Msds-Editor.ps1                 # 启动 GUI
.\Msds-Editor.ps1 cli 文件.docx   # 命令行解析并打印 16 节
.\Msds-Editor.ps1 extract 文件.docx [--query 词] [--scope label|value|all|section] [--json|--tsv] [--out 文件] [--sections 1,3,9]
.\Msds-Editor.ps1 test            # 运行回归测试
.\Msds-Editor.ps1 doctor          # 环境自检
```

双击 `Msds-Editor.bat` 亦可启动 GUI。

依赖: `python-docx` + `tkinter`(标准库)
```bash
pip install -r requirements.txt
```

## 界面功能

| 区域 | 功能 |
|---|---|
| 工具栏 | 固定模板 / 导入产品 MSDS / S1-S3-S9 推断引擎 / 导出 JSON / 17节覆写表单 / **检索框** |
| 左侧目录 | 17 节三级父子级导航树: S0~S16 → 大标题/小标题 → 字段 |
| 右侧表格 | 当前节的 17 节表格视图, 占据右侧全部空间 (三列: 徽章\|序号\|标签\|字段) |
| 检索框 | 输入关键词 → 过滤目录树 (命中节/字段保留), 右侧表格跟随; ✕ 清空恢复完整树 |
| 状态栏 | 当前文件解析统计与告警 |

- **固定模板已锁定** → GUI 不允许替换覆写模板，只使用批准的冠志中文模板
- **导入产品 MSDS** → 读取并显示产品结构，显示源切为产品；不改变覆写模板和编辑权限
- **恢复默认模板** → 重新读取固定模板，显示源切回模板
- **导出 JSON** → 导出当前显示源的 16 节结构化数据，`editable` 由固定白名单计算
- **17节覆写表单** → 在 GUI 内编辑白名单字段、确认覆写并产出 Word
- **推断引擎** → 左侧填写 S1/S3/S9，点击「推导」后在右侧选择目前已开放的 Section 目录（当前为 S2），查看分类、象形图、信号词、H/P 说明、判断条件和人工复核原因；结果只读，不自动写入正式库

### 三列表格: 序号 | 标签 | 字段

```
序号    🔒标签             字段
─────────────────────────────────────────────
8.1    🔒暴露控制
       🔒GHS分类          根据GHS不属于危险物
       🔒氟化橡胶 –FKM     厚度≧0.4mm；穿透时间≧480min.
9.1    🔒外观             乳白色液体
9.3    🔒pH值（1%水溶液）  7-9
```

- **序号列** 固定不可编辑, 窄列
- **标签列** 固定不可编辑 (深色加粗 + 🔒)
- **字段列** 带固定的可编辑/不可编辑徽章；徽章不可点击，权限由 `core/edit_policy.py` 白名单统一决定
- 多行通栏自动配对; 序号自动识别 (`split_seq`), 页码/法规号不被误拆

> 术语约定: **序号** = 编号 (如 11.5), **标签** = 字段名 (如 致突变性),
> **字段** = 内容文本, **父子级** = 节 → 大标题/小标题 → 标签 → 字段 的层级
> (详见 docs/检索系统术语规范.md).

### Section 0 = 页眉/页脚字段

页眉/页脚段落 + 表格全部字段化纳入 **section 0**, 拆成完整父子级:

```
0.页眉页脚
├─ 0.1 页眉
│   ├─ 物料安全数据表 (固定标题)
│   ├─ Version = 1.0
│   └─ 产品名称 = BL-8128W (默认可编辑)
└─ 0.2 页脚
    ├─ 公司名称 = ... (不可编辑)
    ├─ 产品型号 = ... (不可编辑)
    ├─ 修订日期 = ... (不可编辑)
    └─ 页码 = 5 / 5 (不可编辑)
```

## 分层检索与内容提取 (core/extract.py)

把读取结果展开为 **三级父子级树 (节 → 大标题/小标题 → 字段)**, 与 GUI 表格完全对应,
支持单文件/批量检索与提取:

```bash
# 三级父子级输出 (单文件, 默认)
python main.py --extract 文件.docx --sections 1,8,9
# 按关键字检索 (scope: label 标签 / value 字段 / all 全部 / section 节号)
python main.py --extract 文件.docx --query 供应商 --scope label
# 导出 JSON / TSV (TSV 带 BOM, 可直接 Excel 打开)
python main.py --extract 文件.docx --sections 3 --json --out s3.json
python main.py --extract 文件.docx --sections 1,3,9 --tsv --out 提取表.tsv
# 批量多文件统一提取指定节
python main.py --extract 模板.docx PU-1034.docx --sections 1,3,9 --query 供应商 --scope label
```

批量检索入口:

```bash
python 批量化读取/batch_read.py <docx或目录或通配符...> --query 词 --hits --summary
python 批量化读取/batch-read.ps1 <docx或目录> --query "二丙二醇" --hits --summary
```

**代码 API** (`from core.extract import ...`):

| 函数 | 说明 |
|---|---|
| `build_hierarchy(result)` | ParseResult → 三级父子级树 (`list[SectionNode]`) |
| `search_tree(nodes, query, scope)` | 树上检索, 保留父子关系 (二级命中带整棵子树, 三级命中保留该字段) |
| `search_fields(entries, query, scope)` | 扁平检索 (label/value/all/section, 空格多词 AND) |
| `get_field(entries, section, label)` | 精确定位某节某字段值 |
| `extract_many(paths, query, sections)` | 批量处理多文件 → `{文件名: [条目]}` |
| `render_text / render_json / render_tsv` | 文本 / JSON / TSV 输出 |

## 读取器特性 (core/docx_reader.py)

- **页眉/页脚**: 段落 + 表格全读取
- **换行内容**: 多段落保留 `\n`, 软换行 `<w:br/>` 转 `\n`
- **S8 一行跨多列**: 标签格含制表符时自动拆分
- **单列表格内嵌成分表**: S3 单列表格按 `\t`/空格分列、换行分行解析
- **S9 一行/行内多编号拆分**: 标签格内多行各为独立编号字段时按行拆分
- **原始行序保持**: `SectionData.order` 记录文档解析顺序
- **自动编号序号恢复** (`_NumberingResolver`): Word 列表自动编号 (`<w:numPr>`) 恢复实际序号
- **加粗 → 标签列归类**: 加粗文本作为标签信号
- **节标题兼容性**: 节号前单字母前缀 (`v1.`/`l1.`) 与英文 SECTION 长标题自动兼容
- **S3 成分兼容性归一化**: 一行多成分拆分 / 全角符号转半角 / 占位 CAS 保留语义

全库 646 文件解析 0 失败, token 级 0 遗漏。

## 项目结构

```
结构读取/
├── main.py              # 入口 (GUI / --cli / --extract)
├── core/
│   ├── docx_reader.py   # 核心读取器 (16 表格 = 16 节, 页眉页脚, 自动编号)
│   ├── structure.py     # 数据模型 + S3 成分兼容性归一化
│   ├── detectors.py     # 行类型识别 (节标题/字段/成分/说明)
│   └── extract.py       # 三级父子级树 + 分层检索提取
├── gui/
│   ├── main_window.py   # 主窗口 (左目录 | 右表格 + 检索框 + 覆写/推断入口)
│   ├── section_tree.py  # 16节导航 + 固定权限表格呈现
│   ├── inference_window.py # S1/S3/S9 → Section 2 推断工作台
│   ├── overwrite_form_window.py # 17节固定模板表单与 Word 产出
│   └── theme.py         # 配色主题
├── templates/
│   └── MSDS_CN 国彩 模板.docx   # 内化默认模板 (字节级一致)
├── 批量化读取/
│   ├── batch_read.py    # 批量检索 CLI
│   └── batch-read.ps1   # PowerShell 封装
├── tests/test_reader.py # 回归测试
└── outputs/             # 导出 JSON 位置
```

## 测试

```bash
python -m pytest tests -q   # 53 项回归测试
```

## 数据库检索（SQLite 标准字段库）

MSDS 解析结果可入库为 SQLite 四表库（`数据库/正式库/Data Base/msds_standard.db`）：
`schema_field` 标准字段字典 / `msds_model` 型号主表 / `msds_field` 明细长表（父子级全保留）/
`msds_wide` Schema 宽表（126 列）。以**产品型号为唯一索引**检索库内已有数据，
与 docx 直读检索（导入→读取→显示→检索）相互独立。

```bash
python tools/build_msds_db.py <db> <docx或目录...> [--from-xlsx 透视表.xlsx...]  # 入库
python tools/build_msds_db.py <db> --model PEA-4139 [--json|--tsv] [--sections 1,9]  # 按型号检索
python tools/build_msds_db.py <db> --model-search 关键词 [--sections 1,9]            # 关键词检索
python tools/build_msds_db.py <db> --models                                          # 型号列表
```

GUI 工具栏「📚 数据库检索」打开库检索窗口：型号列表点选进入详查，
右侧以与 docx 模式相同的「徽章|序号|标签|字段」四列表格渲染（S11 国标大类归并、
总结句类型化槽位、S8 生物限值子表）。结构基线见 `docs/模板17节结构清单.md`，
需求与验收见 `docs/PRD_数据库检索.md`。

## 17 节固定模板覆写表单（GUI 内完成修改与产出）

工具栏「📝 17节覆写表单」打开固定模板窗口。窗口显示 S0~S16 的完整父子级目录：白名单字段提供输入控件，节标题、父级结构行、标签/序号、子表表头、页码域和其它未获授权字段显示为灰色不可编辑。

权限由 `core/edit_policy.py` 统一控制，不能通过 GUI 徽章、导出 JSON 或临时前端状态放开。S3 成分数据行、S8.2 生物限值数据行和 S15 法规条目分别使用专用动态表格；表头和结构行保持锁定。

典型流程：

```text
固定模板 → 17 节表单 → 字段权限校验 → 用户确认 → pending.docx → 审计 → 正式 Word → 回读校验
```

窗口按钮：

- `保存写入项 JSON`：只保存白名单字段变更，未修改字段保留模板原值。
- `确认覆写并产出 Word`：显示待写入节摘要，确认后调用 `msds_overwrite_engine.py`，完成 pending 产出、字段归属审计和 `verify_output()`。

相关文件：`core/edit_policy.py`（编辑白名单）、`gui/overwrite_form_window.py`（GUI 表单）、`05-覆写模块/msds_form_overwrite.py`（写入项构建）。
