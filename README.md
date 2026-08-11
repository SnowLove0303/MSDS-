# MSDS 结构读取程序

以 `MSDS_CN 国彩 模板.docx` 为标准模板, 提供 **导入 → 读取 → 显示 → 覆写指向** 四段式流程,
用于判断模板覆写指向内容 (哪些模板字段应被新产品数据覆盖 / 保留 / 清空 / 新增)。

## 运行

### 统一指令 Msds Editor（推荐）

已封装为统一指令 `Msds Editor`, 自动定位 Anaconda Python 环境, 无需手动进目录:

```powershell
.\Msds-Editor.ps1                 # 启动可视化界面 (GUI)
.\Msds-Editor.ps1 gui             # 同上
.\Msds-Editor.ps1 cli 文件.docx   # 命令行解析并打印 16 节
.\Msds-Editor.ps1 extract 文件.docx [--query 词] [--scope label|value|all|section] [--json|--tsv] [--out 文件] [--sections 1,3,9]   # 分层检索提取
.\Msds-Editor.ps1 test            # 运行回归测试 (pytest)
.\Msds-Editor.ps1 doctor          # 环境自检 (python / python-docx / tkinter)
```

双击 `Msds-Editor.bat` 亦可启动 GUI。也可通过统一启动器调用:

```powershell
.\统一启动器\System\wl.bat start msdseditor   # 启动 Msds Editor
.\统一启动器\System\wl.bat stop  msdseditor   # 停止 Msds Editor
```

### 直接运行

```bash
cd 结构读取
python main.py              # 启动可视化界面
python main.py --cli 文件.docx   # 命令行解析并打印 16 节
```

依赖: `python-docx` + `tkinter`(标准库)
```bash
pip install -r requirements.txt
```

## 分层检索与内容提取 (core/extract.py)

从代码层把读取结果展开为 **三级父子级树 (节 → 大标题/小标题 → 字段)**,
与 GUI「导入-读取-显示」页眉呈现的表格化内容完全对应, 支持检索与大批量 MSDS 内容提取。

```bash
# 三级父子级输出 (单文件, 默认) — 节 → 大标题 → 字段
python main.py --extract 文件.docx --sections 1,8,9
# --flat: 旧版扁平分层输出 (section → 大标题 → 小标题 → 字段+内容)
python main.py --extract 文件.docx --flat
# 按关键字检索 (scope: label 标签 / value 内容 / all 全部 / section 节号)
python main.py --extract 文件.docx --query 供应商 --scope label
# 导出 JSON / TSV (TSV 带 BOM, 可直接 Excel 打开)
python main.py --extract 文件.docx --sections 3 --json --out s3.json
python main.py --extract 文件.docx --sections 1,3,9 --tsv --out 提取表.tsv
# 批量多文件统一提取指定节 (每文件独立三级树)
python main.py --extract 模板.docx PU-1034.docx --sections 1,3,9 --query 供应商 --scope label
```

分层文本输出示例:

```
[8] 8.接触控制/个人防护
    ─ 8.1 暴露控制
        呼吸系统防护: 喷涂过程中要求有呼吸防护设备。
        手部防护: 喷涂过程中要求有呼吸防护设备。 / 建议戴上防护手套。
        氟化橡胶 –FKM: 厚度≧0.4mm；穿透时间≧480min.
[3] 3. 成分/组成资料
    产品类型: 混合物
        ∟ 聚氨酯聚合物  [CAS: 商业机密 | 含量: 30.0~40.0]
        ∟ 去离子水  [CAS: 7732-18-5 | 含量: >60]
        ∟ 三乙胺  [CAS: 121-44-8 | 含量: 0.1-1]
```

**三级父子级树模型** (`build_hierarchy`):

与 GUI 三列表格 (序号|标题|内容) 完全对应 —— 左侧导航树、覆写指向、检索输出共用同一模型:

```
SectionNode (节, 16 节 + S0 页眉页脚)
 └─ BigTitleNode (二级: 序号子标题如 "8.1 暴露控制", 或带 seq 的字段行如 "9.1 外观")
     └─ FieldNode (三级: 无 seq 字段如 "呼吸系统防护", 说明行 note, 成分行)
```

| 节点 | 字段 | 说明 |
|---|---|---|
| `SectionNode` | `number` / `full_title` / `big_titles` / `direct_fields` | 一级节 (无二级归属的三级字段挂 `direct_fields`) |
| `BigTitleNode` | `seq` / `title` / `value` / `kind` / `children` | 二级: `sub` 纯子标题 或 `field` 带值字段行; `full_title()` 拼 seq+title |
| `FieldNode` | `label` / `value` / `kind` / `editable` / `index` | 三级: `field` / `note` / `component` |

**代码 API** (`from core.extract import ...`):

| 函数 | 说明 |
|---|---|
| `build_hierarchy(result)` | ParseResult → 三级父子级树 (`list[SectionNode]`) |
| `search_tree(nodes, query, scope)` | 树上检索, 保留父子关系 (二级命中带整棵子树, 三级命中保留该字段) |
| `render_tree(nodes)` | 三级树文本输出 (`├─ └─` 缩进, CLI 默认格式) |
| `render_tree_json(nodes)` | 三级树嵌套 JSON |
| `flatten_nodes(nodes)` | 三级树 → 扁平 `ExtractedField` 列表 (TSV/旧 API) |
| `extract_doc(result)` | 整个 ParseResult → 分层条目列表 (`ExtractedField`) |
| `search_fields(entries, query, scope)` | 扁平检索 (label/value/all/section, 空格多词 AND) |
| `get_field(entries, section, label)` | 精确定位某节某字段值 |
| `extract_many(paths, query, sections)` | 批量处理多文件 → `{文件名: [条目]}` |
| `render_text / render_json / render_tsv` | 扁平文本 / JSON / TSV 输出 |
| `export_tsv(result, path)` | 导出 TSV 文件 (带 BOM, Excel 兼容) |

`ExtractedField` 字段: `section` 节号 / `big_title` 大标题 / `sub_title` 小标题 /
`label` 标签(去序号) / `value` 内容 / `seq` 序号 / `kind` (field|sub|note|component|component_header) / `editable`。

> **成分表头可检索**: S3 成分表头作为 `component_header` 条目纳入提取层
> (与 GUI 表头一致), 检索 `化学品名称`/`CAS编号`/`含量`/`w/w`/`Number` 等均可命中;
> 表头**保存实际识别文本** (`Chemical Name | CAS Number | %（w/w）` 等), 英文表头词也可检索;
> 全库 646 文件 462 个成分表全部覆盖, 不作为数据字段参与 get_field 精确定位。
>
> **全库 token 级对照**: 对 646 个文件逐单元格逐 token (中文 3+ 字 / 英文 3+ 字母 / 数字) 与
> 解析结果归一化比对, **0 遗漏** (修复前 43 文件在 S3 遗漏, 根因: 单列表格被误判为表头整格丢弃 +
> 英文表头未识别); 解析失败 0, 成分表全部有成分, 缺失节仅 3 个原文档残缺文件。

## 界面功能

| 区域 | 功能 |
|---|---|
| 工具栏 | 导入模板 / 导入产品 MSDS / 覆写指向分析 / 导出 JSON / **恢复默认模板** |
| 左侧导航 | 16 节三级父子级树: 节 → 大标题/小标题 → 字段 (点击任一级定位到对应节) |
| 16 节结构 | 与 Word 模板表格对齐的节视图: 节标题行 + **三列表格 (徽章 \| 序号 \| 🔒标题 \| 内容)** + 通栏行 + S3 成分表; 点击徽章手动标注内容列可编辑性 |

> **表格内外框**: 整节渲染为一张连续表格 —— **外框 2px 深蓝** (表格最外边界), **内框 1px 浅灰** (行/列分隔线, 连续贯通不割裂), 斑马纹行背景。单元格文本 `sticky=nsew`, **高度随内容自动撑开** (长文本换行后整行变高), **宽度随窗口自适应** (内容列弹性扩展, 换行宽度随画布宽度实时更新, 窗口拖宽/拖窄内容自动重排不溢出)。
>
> **列宽自适应**: 序号列固定窄列; **标题列** ~35% 可用宽 + wraplength 自动换行 (长标题如 `11.8 特异性靶器官系统毒性（一次接触/反复接触）`、`以下为二乙二醇单丁醚…的毒理学参考数据` 自动换行显示全, 不截断); **内容列** ~60% 弹性吃满剩余空间。窗口拖宽 → 标题列/内容列同时变宽; 拖窄 → 自动重排换行。
>
> **S3 成分表内外框**: 与主表格统一 —— 表头为深蓝 cell (化学品名称 | CAS编号 | 含量%), 每个成分行也是独立 cell (徽章 | 名称 | CAS | 含量), 列间 1px 内框竖线 + 行间横线连续贯通, 斑马纹交替。成分行顺序与文档一致。
| 页眉/页脚条 | 顶部摘要显示页眉/页脚 (含页眉表产品名、页脚表公司/修订日期/页号) |
| 覆写指向分析 | 逐字段: 模板现值 vs 产品提供值 → 覆写动作 + 原因 |
| 恢复默认模板 | **真实功能**: 重新读取内化模板副本覆盖比对基准, 显示源切回模板, 字段权限标注复位; 已导入产品保留 (比对仍可用) |
| 状态栏 | 当前文件解析统计与告警 |

> **显示源**: 界面显示由显示源状态控制 —— 启动/导入模板 → 显示模板; 导入产品 → 显示产品;
> 点「恢复默认模板」→ 切回内化模板并重置比对基准。导出 JSON 始终导出**当前显示源**的内容。

### 三列表格: 序号 | 标题 | 内容

凡是有「序号/加粗标题 + 内容」对应关系的行, 一律拆成三列:

```
序号    🔒标题             内容
─────────────────────────────────────────────
8.1    🔒暴露控制
       🔒GHS分类          根据GHS不属于危险物
       🔒氟化橡胶 –FKM     厚度≧0.4mm；穿透时间≧480min.
       🔒符合下列法规要求   危险化学品安全管理条例…\nGB/T 16483…
9.1    🔒外观             乳白色液体
9.3    🔒pH值（1%水溶液）  7-9
```

- **序号列** (8.1 / 9.1 / 9.23): 固定不可覆写, 窄列
- **标题列** (暴露控制 / GHS分类): 固定不可覆写 (深色加粗 + 🔒)
- **内容列**: 带 可编辑/不可编辑 徽章, 徽章只控制内容列 → 覆写模块只替换内容列, **序号/标题永远不被覆写**
- 多行通栏 (含 `\n` 或跨行) 自动配对: "标题：" 与其下内容行合成一列内容
- 序号自动识别: `split_seq("8.1 暴露控制") → ("8.1", "暴露控制")`, 兼容英文标题 (`9.3 pH值`)；页码 `3 / 5`、法规号 `GB/T 16483` 不被误拆
- 配对后的序号/标题列只存在于显示层 (`iter_rows`), **不进覆写比对** → 从源头杜绝被覆写

### 可编辑/不可编辑徽章 (内容列)

| 徽章 | 含义 | 颜色 | 默认 |
|---|---|---|---|
| 🟢 可编辑 | 内容列后续可被覆写替换 (随产品变) | 绿 | 字段/内容默认可编辑 |
| ⚪ 不可编辑 | 内容列固定 (节标题/子标题/页眉页脚固定项) | 灰 | 标题类不可编辑; 页眉/页脚页码、版本、公司、修订日期不可编辑 |
| — 手动标注 | 点击任意徽章可切换该内容列权限, 并持久化到导出 JSON | — | 未标注时用数据模型默认 |

> 权限标注键: `(节, kind, index)` — kind ∈ `field` / `sub` / `note` / `component`。
> 比对后, 模板残留旧值风险 (review) 字段会自动锁定为不可编辑, 防误覆盖; 可手动改回。

### Section 0 = 页眉/页脚字段 (父子级)

页眉/页脚段落 + 表格全部字段化纳入 **section 0**, 并拆成完整父子级:

```
0.页眉页脚
├─ 0.1 页眉 (子标题)
│   ├─ 物料安全数据表 (固定标题, 不可编辑)
│   ├─ Version = 1.0 (不可编辑)
│   └─ 产品名称 = BL-8128W (默认可编辑, 随产品变)
└─ 0.2 页脚 (子标题)
    ├─ 公司名称 = 英德市国彩精细化工有限公司 (不可编辑)
    ├─ 产品型号 = BL-8128W-MSDS (不可编辑)
    ├─ 修订日期 = 2025-3-8 (不可编辑)
    └─ 页码 = 5 / 5 (不可编辑)
```

- **页眉/页脚作为 sub 子标题** (order 交替 + lines 触发, 序号 0.1/0.2), 字段挂其下
- **页脚复合内容拆分**: `英德市国彩精细化工有限公司  BL-8128W-MSDS` → 公司名称 + 产品型号
- **页码识别**: `5 / 5` / `3 / 5` → 独立"页码"字段
- 页眉表格 = **产品名称** (如 PEA-4139), 默认可编辑 (随产品变)
- 页脚固定项 (公司名/修订日期/页码/版本/产品型号) 默认不可编辑
- section 0 不参与自动覆写比对, 其权限完全由手动标注决定
- `产品名称`、`Version`、`修订日期`、`页码` 等字段可直接在界面上标注

### 读取器特性 (core/docx_reader.py)

- **页眉/页脚**: 段落 + 表格全读取 (模板页眉表含产品名 PEA-4139, 页脚表含公司名/修订日期/页号)
- **换行内容**: 多段落保留 `\n`, 软换行 `<w:br/>` 转 `\n`, 值自动换行显示
- **文本格式段落式统一呈现** (`split_text_block`): 部分 MSDS 用 段落+空格+分行标志 (BL-8085 S11) 而非表格格式 — 无冒号短行 (`急性毒性，经口`/`原发性皮肤刺激`/`致癌性`) 识别为**标题行**, 其下内容行配对; 序号子标题 (`11.1 毒理学效应`) 独立为 sub 行; 续行 (`对产品的研究.`/`科学地研究…。`) 并入上一 field 内容; 完整断言句 (`根据GHS不属于危险物`, PU-1034 S2) 与 父标题关系 (`致敏性`→`皮肤致敏性（LLNA）：`) 通过 lookahead 区分, 不误拆 — 所有文本格式均以 Reader 标准三列表格统一呈现
- **S8 一行跨多列**: 标签格含制表符时自动拆分 (如 `手部防护：\t喷涂过程中要求有呼吸防护设备。` → 标签`手部防护` + 值合并)
- **单列表格内嵌成分表**: S3 为单列表格时, 整格文本按 `\t`/空格分列、换行分行解析 (`split_flat_component_text`), 识别表头行后把成分拆为独立行; 兼容列序 浓度|名称|CAS (HPU-7651)、名称/CAS/含量挤一格 (PA-4757)、空格分隔 (OS-8030)、中文 tab (PA-4408); 表头前说明、表头后尾注/S4 溢出通栏保留为说明行
- **无标题行兼容**: 首格空、次列有内容 (如 BL-8085 S5 特殊危害后的 `['', '在着火或爆炸情况下，不要吸进烟尘。']`) → 按原文表格位置界定呈现为 **序号空 | 标题空 | 内容列** 的独立字段, 内容列默认可覆写; 连续空首格行 (PA-4408 S7 跨行续行) 合并为一个块不碎片化; 全库实测 483 个无标题行全部保留
- **S9 一行/行内多编号拆分**: 标签格内多行各为独立编号字段时按行拆分 (英文文件把 `9.19 Dynamic viscosity:\n9.20 Explosion characteristics` 合并进一格); 兼容**行内多编号** — 两编号挤同一行无换行 (`9.19 最低成膜温度MFFT/℃9.20 玻璃化温度Tg/℃：`, RA-15000) 自动断行对齐, 值不错位; 无编号复合标签 (`Combustion value:\nSaturated vapor pressure:`) 不误拆
- **原始行序保持**: `SectionData.order` 记录文档解析顺序 (field/line 交替), `iter_rows` 按原始行序输出 — S8 的 `8.1 暴露控制` 子标题位于呼吸系统防护之前、材质行位于建议之前, 与 Word 表格行序完全一致; 手动构造的 ParseResult 无 order 时用兜底顺序
- **字段留空不告警**: 模板字段留空待填充是常态 (`1.3供应商信息：`/`1.1产品名称：` 子标题行, 全库 646 文件 590 个均有空值), 不再产生 "字段值为空" 告警 — 空值在 GUI 三列表中直观呈现, 不属于解析异常; 告警只聚焦真实问题 (缺失节 / 成分拆分归一化 / S9 真实跳号)
- **自动编号序号恢复** (`_NumberingResolver`): 大量 MSDS 的字段序号 (如 `1.1 产品名称` / `9.1 外观`) 不是文本而是 **Word 列表自动编号** (`<w:numPr>`), `paragraph.text` 不含这些数字 → 序号丢失 (原 644 文件 30021 段受影响)。读取器从 numbering part 解析编号定义 (`numId → abstractNum → lvlText/start`), 按文档顺序计算实际编号并拼接到文本前。lvlText 字面量保留 (`9.%1` → `9.1`, `4.%1.` → `4.1.`), 快照时序 (先取值后递增, 防 9.1 算成 2), 更深级引用浅级计数器、升回浅级重置。修复后 BL-8085 S9 自动编号 `9.1~9.15` 衔接显式 `9.17~9.19`; 全库恢复 29142 个序号字段
- **加粗 → 标题列归类**: 加粗文本作为标题信号 — 无冒号英文标签 (`Waste Disposal Method` / `Empty Container Precautions`, 加粗左列 + 右列有值) 归为**字段行** (标题列受保护不可覆写), 单列加粗长标题 (英文 `Respiratory Protection` 等) 放宽为标题行。排除完整陈述句 (`根据EC指令2006/121/EG,无可用的接触限值信息` 含逗号+长数字 → 内容)。全库 0 token 遗漏
- **英文占位符 CAS**: `Trade secret` / `Business secret` / `Proprietary` 与中文 `商业机密` 同语义, 原样保留不拆词

### 节标题兼容性 (core/detectors.py)

全库实测节标题存在两类变体, 读取器自动兼容:

| 变体 | 规模 | 处理 |
|---|---|---|
| **节号前单字母前缀** `v1.物料及供应商标识` (冠志模板第1节版本标记) / `l1.` (手误) | 233 中文文件 | 剥除前导字母再匹配 → 识别为节1 |
| **英文 SECTION 长标题** `SECTION 1: Identification of the substance/mixture and of the company/undertaking` (81字符) | 38 英文文件 | 英文 SECTION 标题上限放宽至 100 字符 |

修复后: 全库 648 文件 **0 解析失败**, 无 "缺失第1节" (原 147 文件受影响)。仅剩 3 个文档本身缺节会如实报告, 非读取器问题 (经逐文件核对, 缺失内容在原文中确实不存在):

| 文件 | 缺失节 | 原文结构 |
|---|---|---|
| `中文\冠志 guanzhi\PA-4408 msds_CN 冠志.docx` | S8-12 | S7 后直接跳 S13 (表9), 全文无 S8-12 关键词内容 |
| `中文\国彩 guocai\PA-4408 msds_CN 国彩.docx` | S8-12 | 与冠志版结构一致 (S3 表内嵌 S4 溢出通栏) |
| `英文\国彩 guocai\OS-12020 msds_EN Guocai.docx` | S1-2,4-6,8,10-16 | 严重残缺: 仅 3 表格 (S3/S7/S9) + 段落 S14, 其余 60 段落均为空 |

### S3 成分兼容性归一化 (core/structure.py)

全库 646 份 MSDS 的 S3 成分表实测存在三类格式问题, 读取器自动兼容:

| 问题 | 实测规模 | 处理 |
|---|---|---|
| **一行多成分** (名称/CAS/含量三列各含 `\n`, 一个三元组塞 2-6 个成分) | 290 文件 / 602 三元组 | `split_component_cells()` 按 `\n` 拆行对齐 → 拆成 1264 个 `ComponentData` |
| **错用全角符号** (含量 `＞＜～％`, 名称 `－`, CAS `——`) | 188 文件 | 全角→半角 (`＞→>`, `＜→<`, `～→~`, `％→%`, `－→-`), 折叠连续连字符 |
| **占位/无效 CAS** | `商业机密`×303 | 占位符 (`商业机密`/`待确认`/`无单一`) **保留语义**; 纯符号 CAS (`——`/`－`) 置空; 合法 CAS 原样保留 |
| **单列表格内嵌成分表** (整格文本用 `\t`/空格分列、换行分行; 列序可为 浓度\|名称\|CAS) | 39 文件 tab 分隔 (PA-4757/HPU-7651/PA-4408) + 空格分隔 (OS-8030) | `split_flat_component_text()` 识别表头行 (CAS列+名称/含量列) → 表头前/后为说明行, 数据行先抽 CAS (编号/占位符) 再抽含量, 余为名称; 名称/CAS/含量挤一格 (`Acrylic polymer\t9003-01-4 40±1`) 自动拆开; S4 溢出通栏 (含数字含量但无 CAS) 不作成分, 保留为说明行 |
| **英文三列表头** (`Chemical Name \| CAS Number \| %（w/w）` / `Concentration \| Components \| CAS-No.`) | 137 英文文件 | `is_component_header_row` 关键词扩展 + **单行守卫** (表头单元格含 `\n` 或超长即不判表头, 防止整格成分表被误判丢弃) |

归一化原则 (与数据模型同层, 供显示/比对/清单匹配共用):

- `normalize_component_name()`: 全角→半角、压缩空格、折叠连续连字符; **括号是化学名有效部分** (`(DMAC)`/`（BIT）`) 不剥除; 缺连字符 (`N,N二甲基乙醇胺`) 属命名差异, 不强行修复 (由清单归一表处理)
- `normalize_component_cas()`: 去空白、全角→半角; 占位符保留, 纯横线置空, 非法形态 (如 `/` 连写两 CAS) 原样保留不误拆
- `normalize_component_conc()`: 去空白、`＞＜～％` → `><~%`; `±`/文字 (`无数据`) 保留
- 每次拆分/归一化写入 `ParseResult.anomalies` (如 `S3 成分拆分 3 行` / `S3 全角符号归一 2 行`), 改动可追溯

读取器只读不改原文件, 输出即为归一化后的干净数据; 与 `成分清单 (2).xlsx` 比对后 CAS 精确匹配 71.6% + 名称匹配 25.6% (合计 97.2%), 未匹配 10 种均为清单缺项/待人工核对项 (如 `羟基丙烯酸酯聚合物`×22、`矿物油`、`噁唑啉改性丙烯酸类聚合物`)。

## 内化默认模板 + 覆写基座 (core/overwrite.py)

以 `MSDS_CN 国彩 模板.docx` (中文/国彩 MSDS 默认模板) 为覆写基准, **模板已内化进程序**:

- **内化方式**: 模板二进制复制进 `templates/MSDS_CN 国彩 模板.docx`, 与外部源文件 **字节级一致** (SHA256 校验, 测试强制断言), 保证格式零丢失
- **路径解析**: `TEMPLATE_PATH` 优先指向内化副本, 外部源路径作回退 (离线/迁移可用)

**覆写基座** (`overwrite_doc`) —— 只允许 **添加 / 删减 / 编辑** 部分内容, 且 **覆写内容继承原文格式**:

| 操作 | API | 格式保障 |
|---|---|---|
| 编辑单元格 | `set_cell_text(cell, text)` | **逐 run 覆写**, 保留首个 run 的字体/字号/粗细; 多段落 `\n` 保持段落数; 绝不 `cell.text=` (会清空格式) |
| 添加行 | `add_table_row(table, index, template_row)` | `copy.deepcopy` 复制模板行 XML, 新行与模板行格式/列数一致 |
| 删减行 | `delete_table_row(table, index)` | 删除整行 |
| 整文档覆写 | `overwrite_doc(src, changes, out)` | `changes: {(节号, 行位置, 列号): 新文本}`; 行位置按节自动解释 — **成分表节=成分索引** (自动跳过表头/产品类型/列标题), **普通节=表格原始行号** |
| 模板副本 | `copy_template(out)` | 二进制复制内化模板 (覆写前做副本) |

覆写示例:

```python
from core.overwrite import copy_template, overwrite_doc
copy_template("输出.docx")
overwrite_doc("输出.docx", {
    (1, 2, 1): "水性羟基聚酯-丙烯酸分散体 PEA-9999",  # S1 普通节字段行 (行号2)
    (3, 0, 1): "123-45-6",                              # S3 成分0 CAS (成分索引)
    (3, 0, 2): "30-40",                                 # S3 成分0 含量
}, "输出.docx")
```

**格式零丢失验证** (测试断言): 覆写前后单元格 run 格式 (字号/字体/粗体/斜体) 完全一致; 未覆写内容 (页眉/其他节) 原样保留; 添加行格式 == 模板行格式。

## 覆写指向规则 (core/compare.py)

| 指向 | 含义 | 颜色 |
|---|---|---|
| `template` | 模板与产品值一致, 保留模板 | 灰 |
| `product` | 产品提供值 → 用产品覆盖模板 | 绿 |
| `clear` | 模板空字段, 产品未填 → 保持空 | 灰 |
| `add` | 产品新增字段 → 写入 | 蓝 |
| `review` | 模板有旧值但产品未提供 → 残留风险, 需人工确认 | 橙 |

输入集 A (S1/S3/S9) 与外推集 B (S2/S4-16) 分开统计覆盖情况。

## 项目结构

```
结构读取/
├── main.py              # 入口 (GUI / --cli / --extract)
├── core/
│   ├── structure.py     # 数据模型 + S3 成分兼容性归一化层 (拆分/全角转半角/CAS占位)
│   ├── detectors.py     # 行类型识别 (节标题/字段/成分/说明)
│   ├── docx_reader.py   # 核心读取器 (16 表格 = 16 节)
│   ├── compare.py       # 模板 vs 产品 覆写指向分析
│   ├── extract.py       # 三级父子级树模型 + 分层检索提取 (节→大标题→小标题→字段)
│   └── overwrite.py     # 覆写基座 (格式零丢失覆写: 编辑/添加/删减, 成分索引)
├── templates/
│   └── MSDS_CN 国彩 模板.docx   # 内化默认模板 (字节级一致, 覆写基准)
├── gui/
│   ├── theme.py         # 配色主题
│   ├── section_tree.py  # 16节导航 + 结构显示 (含值编辑回调)
│   ├── compare_view.py  # 覆写指向比对表
│   ├── db_view.py       # MSDS 总库 Tab (类目树/回读/编辑/导入/待确认池)
│   └── main_window.py   # 主窗口
├── 数据库正式库/        # 正式总库文件位置 (msds_total.db)
├── db/                  # MSDS 总库 (SQLite)
│   ├── schema.py        # 7表结构 + SECTION_TITLES + connect()
│   ├── catalog.py       # 写入/回读/宽表/检索/增删改 (核心)
│   ├── check.py         # 导入检查报告 (缺失节/必填/异常/重复/未归并)
│   ├── stddict.py       # 半自动标准字段字典 (扫描/归组/频次裁决/待确认池)
│   ├── export.py        # 全库宽表 Excel 导出
│   ├── export_single.py # 单文件 → 数据库格式 Excel (宽表跳空列 + 明细层级)
│   └── audit.py         # 大批量入库就绪度评判 (四层十六项 + 库逻辑 + 批量评估)
├── 批量化读取/          # 批量读取 (build_db.py 静态归一词典共用)
├── tests/test_reader.py # 回归测试
└── outputs/             # 导出 JSON 位置
```

## 测试

```bash
python -m pytest tests -q   # 57 项回归测试
```

## MSDS 总库 (db/ + 主窗口「MSDS 总库」Tab)

### 流程闭环
`导入 MSDS → 检查报告 → 确认添加 → 写入总库 → 类目树回读呈现 → 双击字段编辑 → 写库 → 重建字典(半自动标准字段) → 待确认池裁决`

### 三层形态 (双轨制)
1. **明细 EAV** (`fields`): 全量行级入库, `field_key`(标准) + `raw_label`(原始表达) 双存 —— 回读呈现/编辑用原始表达, 与源文档一致
2. **物化宽表** (`materialize_wide` / `db/export.py`): 列=纯标准字段名(`S{节}.{字段}`), 行=型号, 用于检索/对比/导出
3. **检索** (`search_products`): 标准字段与原始表达双轨命中; 支持 label/value/section/model/成分

### 标准字段机制
- 写入: `resolve_field_key` 按节查 `field_dict` 字典 → 静态 `build_db._NORM` 兜底 → 原样最后兜底
- 归并只发生在**列**; 原始表达永远保留在明细行, 检索两边都认
- 半自动 (P2): `stddict.scan_db`(从库内字段统计) → 归一核归组 → 频次裁决(组内最高表达升为标准名, 其余为别名) → 未命中落 `pending` 待确认池 → GUI 一键采纳(`suggest_merges` 相似建议) 或设为独立标准

### 数据完整性 (禁止缺失)
- 所有 fields 行 (含 S0 页眉页脚 / 节标题 / sub / note / field) 全量入库
- components + pictograms(blob) 完整入库
- 重复检测: sha256 相同 → `exists_sha`; 同类目下型号相同 → `exists_model` (均不重复写入)
- 删除为软删除 (active=0); revisions 记录 add/delete 版本历史

### GUI 用法 (主窗口 → 「MSDS 总库」Tab)
- 新建库/打开库; 启动时自动打开 `结构读取/数据库正式库/msds_total.db`
- **导入 MSDS**: 选 .docx → 检查报告弹窗(缺失节/必填/解析异常/未归并/重复) → 填一级类目+型号 → 添加到总库
- **左侧视图切换**: `📂 类目树` (📁 产品类型 → 型号) ⇄ `📋 型号目录` (以产品型号为索引平铺全部型号, 列=型号|类目|语言|产品名)
- **型号目录点击加载**: 点任一型号行 → 右侧复用 16 节导航 + 表格呈现该型号全部信息 (回读自总库)
- 检索框双轨检索 (标准字段/原始表达/内容), 两种视图均按命中型号过滤
- 右侧型号详情: 16 节导航 + 表格呈现; **双击值单元格** → 弹窗编辑 → 写库
- **重建字典**: 从库内全部字段重新统计并重建标准字段字典 (active 保留作种子, pending 重算)
- **待确认池**: 列出未归并标签与相似标准字段建议; 「采纳建议」并入标准字段(保留原表达为别名) / 「设为独立标准」
- **导出宽表**: 行=型号, 列=标准字段+成分, Excel

### 数据层 API (无 GUI 亦可脚本调用)
```python
from db.schema import connect
from db import catalog, stddict, check, export
conn = connect("总库.db")
pid, status = catalog.add_product(conn, result, category="PU", model_name="PU-1000")
r = catalog.rebuild_product(conn, pid)          # 回读为 ParseResult
hits = catalog.search_products(conn, "毒性")     # 双轨检索
info = export.export_wide(conn, "宽表.xlsx")     # 导出宽表
rpt = check.check_doc(result, conn)             # 导入检查报告
res = stddict.build_field_dict(conn, stats=stddict.scan_db(conn))  # 重建字典
```

## 单文件数据库格式 Excel (db/export_single.py)

对单份 MSDS 生成"数据库格式"Excel (双 Sheet), 与总库宽表同构:

- **Sheet1 数据库宽表**: 行=1, 列 = 型号/一级类目/产品名称/版本/修订/页眉/页脚
  + 数据列**严格按节 1→16 升序** + 节内首现行序
  + 非空字段列 (`S{节}.{标准字段}`); **非空说明行也入列** (`S{节}.说明{n}`),
    纯 note 节 (如 S16 免责声明) 不再整节消失
  + **成分列插在 S3 节内** (`S3.成分{i}名称/CAS/含量`, 紧跟 S3 字段之后, 而非排到所有节末尾)
  + 空值字段列**跳过**, 避免大量空列混乱
- **Sheet2 完整明细**: 逐行全量 (禁止缺失), 九列结构:
  `节 | 行类型 | 序号 | 原始标签 | 标准字段 | 值 | 父级 | 可编辑 | 备注`
  - **序号列**: 保留原文序号 (`row.seq`, 如 `9.1~9.19`), 恢复原始顺序与上下级
  - **父级列**: 节→子标题→字段 父子链 (栈式, 同级 sub 并列, 分组标题吸收其下子项,
    并列防护类别如眼睛/身体防护自动回到分组父级)
  - **备注列**: 分组标题(引导其下子项)/空值待填/孤儿(无标签挂接前项) 标注移出值列,
    **值列只放 MSDS 真实内容** (不混入人工占位文字)
  - **成分拆行**: 每成分拆 名称/CAS/含量 三行, 与宽表成分列一一对应
  - 节标题(深蓝通栏) / 子标题 sub(浅蓝加粗) / 字段 / 说明 note(斜体) / 分组标题 / 成分

```python
from db.export_single import export_single
info = export_single(result, conn, "BL-8128W_数据库格式.xlsx",
                     category="测试类", model_name="BL-8128W")
```

### 解析修复 (BL-8128W 全库核对发现)
- `core/docx_reader._parse_field_row` else 分支: 2 列空值行 `['文本', '']` join 时
  不再产生 `'文本 / '` 尾部斜杠 (BL-8128W S11 '该产品无可用的毒理学研究。/'
  原被判为畸形字段, 现为说明行 note)

### 解析修复 (全库 646 评判发现)
- **S3 'Mixtures' 声明保留**: 英文模板 S3 单列 `3.1 Mixtures\n· polyether urethane`
  原来 `_is_mix_header` 命中后整格 continue, 'Mixtures' 标题本身被丢弃
  (对应中文 '产品类型：混合物') — 现在作为 sub 标题保留 (iter_rows 自动拆序号),
  成分 note 仍正常延迟提升到 components。修复后 token 级 0 遗漏。
- **复数占位符 CAS 保留**: `Trade Secrets` / `Business Secrets` 原只匹配单数
  `secret`, 复数被去空格成 `TradeSecrets` (占位语义破坏) — 现在
  `_CAS_PLACEHOLDER_EN_SPACE_RE` 支持复数, 原样保留含空格文本。

## 大批量入库就绪度评判 (db/audit.py)

对指定文件夹 (递归扫描 *.docx) 逐文件执行 **写入与输出对照**, 从四个递进层次
验证数据完整性, 并出具数据库逻辑审查 + 大批量能力评估:

```bash
python -m db.audit "F:\...\标准化测试一库"          # 评判文件夹, 报告写到该目录
python -m db.audit "F:\...\MSDS" --out 报告目录      # 指定报告输出位置
```

**四层十六项判定** (每项 PASS / WARN / FAIL):

| 层次 | 检查项 | 说明 |
|---|---|---|
| L1 解析 | 节完整 / token 0遗漏 / 结构 / 序号 | reader → ParseResult 与源 docx 逐 token 比对 |
| L2 写入 | 字段全量 / 成分全量 / key收敛 / raw_label逐行 / token / 表头 / 幂等 | ParseResult → SQLite |
| L3 回读 | 节/内容/顺序/成分/表头 | SQLite → ParseResult 往返一致 |
| L4 输出 | 宽表节序 / 节覆盖 / 明细全量 / 值保真 | SQLite → Excel (export_single) |

**判定规则** (不阻断, 如实分级):
- **FAIL** = 内容/结构损坏 (token 真丢失 / 字段或成分未入库 / 回读不一致 / 导出缺内容节)
  — 阻断大批量, 需修复
- **WARN** = 原文结构特性 (原文无页眉页脚 / 原文缺节残缺文档 / 原文空节 /
  自动编号恢复的序号怪癖) — 记录但不阻断, 报告中列出待核查清单
- **token 子序列放宽**: S2 P 代码连写标签 (`存储存储已锁定。`) 被 reader 拆成
  标签+值 (`存储` / `存储已锁定。`), 整串 token 不再连续但字符仍在 → 子序列
  匹配即视为已覆盖, 不算遗漏 (信息未丢失)

**输出**: `audit_report.json` (结构化) + `audit_report.md` (人读报告),
含数据库逻辑审查 (EAV 规范化 / 外键级联 / 唯一约束 / 索引 / 软删除) 与
646 文件外推吞吐估算。

**全库实测** (646 文件, 2026-08-11): 四层 646/646 通过, **0 FAIL / 0 ERROR**,
`ready_for_bulk = true`, 约 6 分钟全库 (单文件均值 <0.6s)。597 全项通过 +
49 个 WARN 全部为原文特性 (16 序号怪癖+无页眉页脚 / 15 序号倒序 / 8 无页眉页脚 /
4 原文空节 / 2 残缺文档 / 2+2 混合)。表头持久化后 `component_header` 随字段完整入库并回读。

### ChatGPT 批量入库 12 项标准 → 实现对照

依据 ChatGPT 集中测试分析 (12 项批量入库标准) 的实施覆盖:

| # | 标准 | 级别 | 实现 |
|---|---|---|---|
| 1 | 节完整性 (16 节全解析) | A | L1 `audit_parse`: sections_total 与原文 16 节逐节核对 |
| 2 | 内容覆盖 (token 0 遗漏) | A | `_TOKEN_RE` 逐 token 比对 + 子序列放宽 (`_is_subseq`) |
| 3 | 节顺序 (1..16 单调) | A | L1 节序 + L4 宽表节序单调性 `section_seq_monotonic` |
| 4 | 父级归属 (节→分组→实体) | A | 父级链栈 (sub/分组压栈) + `parent_id` 字段显式记录 |
| 5 | 唯一 node_id | A | `document_tree.py` S{n}/S{n}.{seq}/GROUP/F/NOTE/ING; 重复 seq 追加 `.cnt` 计数, 全库 0 重复 |
| 6 | 节点类型标注 | A | 每节点 `type`: section/subsection/group/field/text/entity |
| 7 | 可重复实体拆分 | B | 成分 → `entities` (S3.ING.{i}.name/cas/conc), 不再挤进 Value |
| 8 | 完整上下文路径 | B | node_id 自描述 + `parent_id` 链 (节→分组→实体) |
| 9 | Raw-Normalized 分离 | B | 字段 `value`+`raw_value` 双轨; 成分 `name/cas/conc` + `raw_name/raw_cas/raw_conc`; L2/L3 全量验证, 仅 value 非空而 raw 空才判丢失 |
| 10 | Schema 冻结 | C | `schema.py` 幂等建表 + 迁移 (`_COMP_RAW_MIGRATIONS`); `audit_schema` 审查 EAV/外键/唯一/软删除 |
| 11 | 机器-人类报告分离 | C | `document_tree.json` (机器树) 与 `audit_report.md` (人读) 分文件输出 |
| 12 | Round-trip 测试 | C | L2 写 → L3 回读 (`rebuild_product`) 往返一致; L4 再导出宽表验证 |

A 级 (数据正确性) 全量通过; B 级 (结构化) 由 document_tree 校验; C 级 (工程规范)
已冻结。剩余一项待办: 让 `export_document_tree` 输出接入主流程 (当前由
`audit_rebuild` 内存校验, JSON 导出为独立入口)。
