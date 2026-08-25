# MSDS 推导产出系统（冠志）

本目录是**一个系统**的工作区：围绕 MSDS（化学品安全技术说明书）的
**表单录入 → 结构检索 → 标准入库 → 内容推断 → 模板覆写 → 多格式输出**
全生命周期，各模块已按职责分区（编号即流水线顺序），模块间通过标准 JSON
契约（write_items / 标准范式 / 覆写引擎输入契约）衔接。

> 归档区镜像见 `../MSDS--TDS-repo`（由 `_scripts/_repo_deploy.ps1` 部署生成，
> 只推代码/文档/模板/结构化库表，不推产品文档与法规原文 PDF）。

## 目录结构

```
MSDS\
├── 01-表单系统\      表单窗口产物与测试：表单系统与推断引擎（S2 书写规范）、表单-覆写测试
│                       （表单代码 gui/form_window.py + core/form_schema.py 在 02-检索系统 内运行，
│                         部署时复制到归档区 01）
├── 02-检索系统\      MSDS 结构读取/检索系统（独立 git 仓库）：
│                       core（docx_reader/extract/msds_db/schema…）、gui、tools、docs、
│                       tests、批量化读取、main.py / Msds-Editor.ps1
├── 03-数据库\        数据库区：
│                       正式库（Data Base/msds_standard.db、DB、标准字段数据库Excel）、
│                       测试库、MSDS（原始中英文文档）、数据清理模块（msds_extract + 标准模板）
├── 04-推断引擎\      内容推断：判断skill（msds-inference-write：reader→build→docx）、
│                       法规匹配库（01~08 法规/标准分类归档 + 推断引擎数据）
├── 05-覆写模块\      模板覆写：msds_overwrite_engine.py、run_db_overwrite.py、
│                       write_items_*（当前版 *_pure_v3.json）、SKILL、outputs
├── 06-联网搜证\      联网抓取/搜证工具（法规匹配库 tools 的核心工具集）
├── 07-多格式输出\    检索/导出产物（原 结构读取/outputs）
├── docs\             系统文档（PRD 等）
└── _scripts\         部署/运维脚本（_repo_deploy.ps1 生成归档区）
```

## 数据流（模块衔接）

```
01 表单录入 ──写入项(S1/S3/S9)──▶ 02 检索读取 ──标准范式──▶ 03 入库(SQLite)
     │                                                      │
     │                                                      ▼
     └────────◀── 05 覆写套模板 ◀── 04 推断补齐(S2/S4~16) ◀── 检索 write_items
                        │
                        ▼
                 07 多格式输出（docx / xlsx / tsv / json）
```

- **表单→覆写契约**：`core/form_schema.py`（02-检索系统）产出写入项，`05-覆写模块/field_maps*.json` 对齐标签。
- **数据库→覆写契约**：`run_db_overwrite.py`（05-覆写模块）从 `03-数据库/正式库/Data Base/msds_standard.db`
  检索型号 → 构建 write_items → 调用覆写引擎。
- **检索→覆写契约**：见 `02-检索系统/docs/MSDS标准检索信息输出格式规范(覆写引擎输入契约).md`。

## 常用入口

| 用途 | 命令 |
|---|---|
| 检索系统 GUI / CLI | `02-检索系统\main.py` 或 `02-检索系统\Msds-Editor.ps1` |
| 批量读取 | `02-检索系统\批量化读取\batch_read.py` |
| 标准库入库 | `02-检索系统\tools\build_msds_db.py` |
| 数据库→覆写流水线 | `05-覆写模块\run_db_overwrite.py --model <型号>` |
| 部署归档区 | `_scripts\_repo_deploy.ps1` |

## 约定

- `_scratch\` = 临时脚本/中间产物（不部署、不进 git）；`_archive\` = 历史版本归档。
- 目录/路径变更以 `_scripts/_repo_deploy.ps1` 为唯一权威，改动后需重新部署归档区。
