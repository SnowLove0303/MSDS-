# MSDS 覆写引擎

以**模板格式为准**，把写入项 JSON 覆写进指定 MSDS 模板：只替换文本，格式继承模板原单元格。
输入是检索/推断产出的**写入项**（节 → 字段 → 值），输出是基于模板覆写出来的新 Word 文档
（**不是源文件拷贝**）。

## 流水线

```
真实 MSDS.docx
   │  结构读取（结构读取/模块，只读检索，17 节）
   ▼
make_real_write_items.py  ──►  写入项 JSON（write_items_*.json）
   │
   ▼  clean_write_items.py（剔除"仅说明"内容，见下）
   ▼
write_items_*_pure.json（纯 MSDS 版写入项）
   │
   ▼
msds_overwrite_engine.py --template 模板.docx --write-items 写入项_pure.json --out 输出.docx
   │
   ▼
输出.docx（模板 B 被覆写；闭环校验通过）
```

**注意**：写入项若来自推导方案 content（含 msds/note 区分），必须先经 `clean_write_items.py`
剔除 note 内容（说明/编辑指令/推导理由），否则会把"建议设置专用公共邮箱""说明：本品各健康
危害成分…"这类不该出现在正式 MSDS 的说明写进输出。纯 MSDS 版写入项 `.json` 就是覆写的输入。

**17 节**：`S0 页眉页脚` + `S1~S16`。S0 通过专用路径覆写页眉/页脚表格
（产品名称/Version/公司名称/产品型号/修订日期），页码为域自动生成不覆写。

- 写入项由 `make_real_write_items.py`（真实 MSDS）或 `make_write_items.py`（测试库 Excel，模拟推断引擎输出）生成；
- 覆写引擎只负责「把指定内容覆写进指定模板的指定位置，采用指定格式」。

## 运行

```bash
# 单个产品覆写
python msds_overwrite_engine.py \
  --template "..\数据库\测试库\PEA-4139 MSDS_CN 冠志 模板.docx" \
  --write-items write_items_os1330.json \
  --out "..\数据库\测试库\输出库\OS-1330 标准化输出.docx" \
  --sections 1,2,3,9

# 从真实 MSDS 生成写入项
python make_real_write_items.py "真实MSDS.docx" --out write_items.json
python make_real_write_items.py <目录> --out <输出目录>        # 批量
```

### CLI 参数

| 参数 | 说明 |
|---|---|
| `--template` | 模板 docx（必填） |
| `--write-items` | 写入项 JSON（必填） |
| `--out` | 输出 docx 路径（必填） |
| `--sections` | 仅处理指定节，如 `1,3,9` |
| `--ref-pick` | 新增行参考行策略：`auto`(默认)/`last`/`label:标签名`/`row:行号` |
| `--keep-structure` | 保留模板结构的节（不删除模板字段），如 `2`；默认含 2 |
| `--empty-policy` | 空值覆写策略：`warn`(默认)/`overwrite`/`preserve` |
| `--field-map` | 字段映射配置 json（默认同目录 `field_maps.json`） |

### note 清洗（纯 MSDS 版写入项）

```bash
python clean_write_items.py write_items_os1330.json      # → write_items_os1330_pure.json
python msds_overwrite_engine.py --template 底版.docx --write-items write_items_os1330_pure.json --out 输出.docx
```

## 写入项 schema

```json
{
  "sections": {
    "1": [ {"seq": "", "label": "中文名称", "value": "水性双组份聚氨酯催干剂 OS-1330"}, ... ],
    "3": { "产品类型": "混合物",
            "components": [ {"name": "水", "cas": "7732-18-5", "conc": ">50"}, ... ] },
    "9": [ {"seq": "9.1", "label": "外观", "value": "乳白色液体"}, ... ]
  },
  "keep_structure": [2],
  "empty_policy": "warn"
}
```

顶层 `keep_structure` / `empty_policy` 可被 CLI 参数覆盖。

## 核心机制

### 格式捕获-重放（格式继承）
捕获目标单元格每个段落的 pPr + 代表 run 的 rPr；新值按 `\n` 拆段逐段重放，
多余段落删除、不足克隆末段格式。超链接/字段域/图片随旧内容清除，**超链接样式
（下划线/蓝色）不继承**，只继承普通文本格式。

### 排版空格围栏
模板很多字段行用空格做对齐（如 `9.2  嗅觉阈值：      `）。覆写时只替换文字/符号，
保留序号后空格与尾部对齐空格；序号为空时沿用参考行序号，冒号跟随参考行风格。

### 字段映射层（打标签/建通道）
以模板为主：把源文件/检索输出的字段写法映射到模板标准字段。规则在 **`field_maps.json`**
外部配置，模板更换只需换配置，不改引擎代码（详见《模板适配指南.md》）。

### keep_structure（保留模板结构节）
`keep_structure` 节（默认 S2 GHS 标准结构）模板字段行**不删除**，只覆写匹配项；
缺字段的行保留模板默认值，但会触发必需字段告警（见下）。

### 空值覆写策略
写入项含该字段但值为空 → 锚定并覆写为空。三种策略：
- `warn`（默认）：覆写清空 + 模板原值非空时告警；
- `overwrite`：静默清空；
- `preserve`：跳过，保留模板原值。

空值覆写会**清除排版空格残留**（否则读回非空，污染闭环校验）。

### 闭环校验
覆写后用结构读取 `read_msds` 读回输出文档，与写入项逐字段对比（值、S3 成分归一化、
保留字段存在性）。CLI 默认执行，失败退出码 1。
对数字开头的化学品名（如 `2-丁氧基乙醇`）登记完整标签，避免被 `norm_field`
误拆序号导致漏匹配。

### S2 象形图

S2 象形图不再使用源文件图片顺序、`[象形图]` 占位符或旧的黑色符号作为事实来源。
统一使用 `04-推断引擎/推断引擎程序/pictogram_registry.py` 管理的 `GHS01~GHS09`
标准代码，并从以下唯一图库取图：

`F:\正式项目与模块化内容\冠志\MSDS\04-推断引擎\判断skill\templates\pictograms`

写入值示例：`GHS05 腐蚀；GHS07 感叹号`。覆写引擎会将代码文本写入模板值格，
同时嵌入对应的红色菱形 PNG；检索 GUI 使用同一注册表显示图片，表单 GUI 使用同一
注册表选择图片。因此 `GHS-象形图` 只有在值中包含 `GHS01~GHS09` 代码时才进入
象形图通道，否则仍按旧源文件的“GHS 标签要素”语义处理，避免误把标签上的有害
成分说明替换成图形。

## 隐患防御机制

| 隐患 | 防御 |
|---|---|
| 结构读取启发式脆弱（源写法变化） | S2 续行并入写入项；覆写侧「新增过多」告警 |
| 映射通道锚定单个模板 | `field_maps.json` 外部配置 + `--field-map` |
| keep_structure 掩盖缺字段 | `REQUIRED_FIELDS` 必需字段缺失告警 |
| 空值清空模板默认值不可审计 | `empty_policy` 三策略 + warn 告警 |
| 超链接格式污染覆写值 | `_capture_para` 优先取普通 run 的 rPr |

## 项目结构

```
覆写引擎/
├── msds_overwrite_engine.py   # 主引擎（覆写 + 闭环校验 + CLI）
├── msds_form_overwrite.py     # 17 节中文控制台表单 → write_items → Word
├── msds-form-overwrite.ps1    # PowerShell 中文交互入口
├── make_real_write_items.py   # 真实 MSDS → 写入项 JSON（批量流水线雏形）
├── make_write_items.py        # 测试库 Excel → 写入项 JSON（模拟推断引擎输出）
├── clean_write_items.py       # 写入项 note 清洗（剔除仅说明内容，生成 _pure 版）
├── field_maps.json            # 字段映射外部配置（默认）
├── write_items_*.json         # 写入项示例（OS-1330 / BEK-750 / test）
├── README.md                  # 本文档
└── 模板适配指南.md             # field_maps 模板适配示例文档
```

## 17 节中文交互表单（Python / PowerShell）

`msds_form_overwrite.py` 将结构读取器、17 节父子级骨架和覆写引擎串成一条控制台流水线。它会从指定模板读取可编辑字段，按 `S0~S16`、父级、序号、标签显示，直接回车保留模板值，输入 `!清空` 显式清空字段，多行内容以单独输入 `结束` 完成。

Python 入口：

```powershell
python msds_form_overwrite.py --template "F:\正式项目与模块化内容\冠志\MSDS\03-数据库\正式库\推导方案\PEA-4139 MSDS_CN 冠志 模板.docx" --sections all
```

PowerShell 中文入口：

```powershell
.\msds-form-overwrite.ps1 -Sections all
.\msds-form-overwrite.ps1 -Sections "1,3,8,9,15" -Out "输出目录\产品.docx"
```

流水线会先保存 `write_items_*.json`，显示待覆写节和字段数量，等待确认后才生成 Word；生成后调用 `verify_output()` 结构回读校验。默认空白输入保留模板值，显式 `!清空` 时才清空对应字段。S3 成分、S8.2 生物限值和 S15 法规条目使用专用子表/动态行输入。

已有 JSON 可用于重复产出或自动化：

```powershell
python msds_form_overwrite.py --input-json write_items_xxx.json --sections all --yes --out "输出.docx"
```

当前覆写引擎仍遵守项目的唯一中文冠志模板约束；`--template` 用于明确指定该模板路径，若传入非批准模板，运行会在覆写前拒绝并提示。

## 依赖

- `python-docx`、`openpyxl`、`lxml`
- 结构读取模块（`F:\正式项目与模块化内容\冠志\MSDS\02-检索系统`，用于「查」与闭环校验）
