---
name: msds-overwrite
description: 用标准 MSDS 模板（PEA-4139 冠志/国彩等）将写入项 JSON 覆写成标准格式的 MSDS 输出文档。用于 MSDS 覆写、安全技术说明书标准化、修复写乱的 MSDS、模板格式继承覆写、写入项生成、闭环校验、格式签名验证、模板适配（field_maps）等任务。当用户要求"用模板重写 MSDS""标准化输出""覆写引擎""PEA-4139 模板""EC-1801/OS-1330 等产品标准化""按 GB/T 16483 重排"时启用。
---

# MSDS 覆写（模板驱动）

把真实/检索得到的 MSDS 内容，按 **标准模板的序号、标签、字段结构、总体格式** 覆写成一份格式标准的新 Word 文档。模板是契约，写入项只负责填值。

## 非破坏性契约（铁律）

- **模板为准**：模板的字段行是骨架，全部保留；写入项只做值源。删除模板行必须写入项显式 `"delete": true`。
- **格式继承**：复用模板行时，标签格（序号 run + 中文 run 的双 run 结构）100% 继承模板原文，只覆写值格。
- **不要误删**：写入项没覆盖的模板字段默认保留（模板驱动）；不要为了"对齐写入项"而删除模板骨架。
- **数据不丢**：写入项来自源 MSDS 的数据必须全部出现在输出（尤其 S12 生态数据的 Log Kow/BCF/BOD/COD/迁移性/PBT 等）。
- **闭环必过**：覆写后 `verify_output` 闭环校验失败即视为未完成，必须修复。
- **路径/命令/字段/代码原样保留**：Windows 路径、CAS、成分名、序号、标签、化学式不擅自改写。

## 核心概念

| 概念 | 说明 |
|---|---|
| 17 节结构 | `S0 页眉页脚` + `S1~S16`（GB/T 16483）。S0 走专用路径覆写页眉/页脚表格，页码是域自动生成不覆写 |
| 模板驱动 | `keep_structure="all"`（默认）：全部节以模板为骨架，模板字段行不删除，只覆写匹配值 |
| 写入项 | JSON：`{"keep_structure":"all","sections":{"1":[{label,value}...],"3":{...},...}}`，label 是匹配键，value 是值源 |
| 字段映射 | `field_maps.json` 把源写法→模板标准写法（建通道）。映射未命中 → 字段会"删+增"而非"改" |
| 标签格双 run | 模板标准：序号 run（Times New Roman 10.5pt 加粗）+ 中文 run（宋体/Arial 12pt 加粗）。S9 个别行（9.2/9.6/9.7）三 run 带尾部对齐空格 |
| 底版 | 模板去产品特有残留后的覆写输入（如 `EC-1801 底版_冠志.docx`） |

## 标准工作流

```
源 MSDS.docx（内容真值）
   │  结构读取只读检索  read_msds
   ▼
make_real_write_items.py ──► 写入项 JSON（write_items_<产品>.json）
   │
   ▼  （如模板含上一步骤产品残留：prep_ec1801_base.py 先做底版）
msds_overwrite_engine.py --template 模板/底版 --write-items 写入项 --out 输出
   │
   ▼
输出.docx（模板格式继承；闭环校验通过）+ 结构读取回读验证
```

### 步骤 1 — 读源 MSDS（结构读取）

```bash
cd "F:\正式项目与模块化内容\冠志\MSDS\02-检索系统"
python main.py --cli "<源MSDS.docx>"                      # 全 16 节摘要 + 告警
python main.py --extract "<源MSDS.docx>" --sections 0,1,3,9,12   # 三级树，只看关键节
python main.py --extract "<源MSDS.docx>" --json --sections 12 --out s12.json
```

**先看告警**（`anomalies`）：写乱的 MSDS（如 S12 生态数据堆砌在单格）需要靠覆写重排，不能照搬。

### 步骤 2 — 生成写入项 JSON

```bash
cd "F:\正式项目与模块化内容\冠志\MSDS\05-覆写模块"
python make_real_write_items.py "<源MSDS.docx>" --out write_items_<产品>.json
python make_real_write_items.py "<目录>" --out "<输出目录>"     # 批量
```

生成后**人工核对**：
- 中文名称补型号规则已应用（`中文名称 + 空格 + 型号`）；
- S3 成分 name/cas/conc 三列完整；
- S12 生态数据无丢失（Log Kow 2.68、BCF、土壤迁移性、PBT/vPvB、BOD/COD 等）；
- 顶层加 `"keep_structure": "all"`。

### 步骤 3 — 模板/底版准备

模板是标准格式基准：
- 模板本身干净 → 直接用模板。
- 模板含**上一步骤产品特有残留**（如 PEA-4139 模板里有二乙二醇单丁醚引导句/手套材质行）→ 先做底版：
  - 参考 `prep_ec1801_base.py`：删除 S8 手套材质行、S11/S12 引导行、S13 欧盟句、S12 重复 12.1，清 S8 手部防护制表符。
  - 产出 `EC-1801 底版_冠志.docx` 放 `数据库\测试库\输出库\`。

### 步骤 4 — 覆写

```bash
cd "F:\正式项目与模块化内容\冠志\MSDS\05-覆写模块"
python msds_overwrite_engine.py \
  --template "<模板或底版.docx>" \
  --write-items write_items_<产品>.json \
  --out "F:\正式项目与模块化内容\冠志\MSDS\03-数据库\测试库\输出库\<产品> 标准化输出.docx"
```

**看日志判定通道命中**：
- ✅ `S<N> 改 row...` = 复用模板行覆写值（格式继承，正确）；
- ⚠️ `S<N> 增 [字段]` = 模板没有该字段，克隆新行插入。若同节大量"增"触发 `S<N>: 写入项 X 个字段中 Y 个靠新增插入...请核对字段映射` → **映射未命中，补 field_maps**。
- 闭环校验失败 → 逐条 `FAIL` 修复后重跑。

### 步骤 5 — 验证（三重）

1. **闭环校验**：引擎已自动执行，`闭环校验通过 OK`。
2. **残留扫描**：输出不得含上一步骤产品关键词（如 PEA-4139 → `二乙二醇单丁醚/乳白色液体/P210/非危险品`）；S3 成分 CAS="商业机密"可能是源真实数据，不误判。
3. **格式签名**：抽查标签格双 run 结构 + 与底版逐行对比（脚本见 `references/verification.md`）。

## 关键路径

| 用途 | 路径 |
|---|---|
| 覆写引擎 | `F:\正式项目与模块化内容\冠志\MSDS\05-覆写模块\msds_overwrite_engine.py` |
| 写入项生成 | `F:\正式项目与模块化内容\冠志\MSDS\05-覆写模块\make_real_write_items.py` |
| 底版预处理 | `F:\正式项目与模块化内容\冠志\MSDS\05-覆写模块\prep_ec1801_base.py` |
| 字段映射 | `F:\正式项目与模块化内容\冠志\MSDS\05-覆写模块\field_maps.json` |
| 结构读取 | `F:\正式项目与模块化内容\冠志\MSDS\02-检索系统\main.py`（GUI/--cli/--extract） |
| 模板基准 | `F:\正式项目与模块化内容\冠志\MSDS\03-数据库\测试库\PEA-4139 MSDS_CN 冠志 模板.docx` |
| 输出库 | `F:\正式项目与模块化内容\冠志\MSDS\03-数据库\测试库\输出库\` |
| 已有写入项 | `write_items_ec1801.json`、`write_items_os1330.json`、`write_items_bek750.json`（参考 schema） |

## 参考文档索引

按需读取，不全部预读：

- `references/commands.md`：完整 CLI 参考（覆写引擎 + 结构读取 + 写入项生成）。
- `references/write-items-schema.md`：写入项 JSON 详解（S0~S16 各节写法、delete、空值策略、keep_structure）。
- `references/troubleshooting.md`：常见问题排查（映射未命中/编号重排/模板残留/标签格降级/空值告警）。
- `references/verification.md`：格式签名验证脚本 + 残留扫描清单。

## 最佳实践

- **字段映射优先于手工补**：新字段写不进模板 → 先查 `field_maps.json`，缺规则就补，别靠 `ref_pick` 硬插。
- **seq 不用刻意写**：模板驱动下复用行序号继承模板；写入项 seq 主要用于新增行。若 S9 编号重排 → 清空写入项 seq 沿用模板编号。
- **空值要谨慎**：模板默认值可能是有意义的法定值（如信号词）。写入项空值默认 `warn`（覆写+告警）；如需保留模板默认值用 `"empty_policy": "preserve"`。
- **多产品复用**：同一模板配多份 `write_items_<产品>.json` + 可选 `field_maps.<模板>.json`，不改引擎代码。
- **改引擎先备份**：`msds_overwrite_engine.py.bak` 保留可回退版本；改完跑一次完整覆写闭环验证。
