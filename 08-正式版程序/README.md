# MSDS 正式版 Web 操作程序（第一阶段）

本目录提供第一阶段 Web 工作台：

- 型号查询：调用 `03-数据库/正式库/Data Base/msds_standard.db`，型号唯一索引，支持型号/字段/值模糊检索。
- 分库总览：Web 页面直接展示中文总型号库、字段映射库、CAS 库和型号大类索引的路径及统计信息。
- 型号归类：读取 `F:\冠志工作空间\产品\TDS MSDS\TDS MSDS\产品 TDS MSDS -- WORD版本` 父级目录，建立“产品大类 → 型号”父子级索引。
- 表单功能：只展示 `core.form_schema` 定义的 S1、S3、S9 可编辑字段；Section 9 仍按冻结的 37 项标准字段工作。切换 Section 时先写入本地草稿并同步保存到 `_codex_work/web_drafts`。
- 覆写功能：使用既有 `05-覆写模块/msds_overwrite_engine.py`，固定中文 Word 模板驱动，执行产出后回读校验。
- 空值策略：表单空值由既有表单契约统一变为 `无数据`。

## 启动

在 PowerShell 中执行：

```powershell
Set-Location 'F:\正式项目与模块化内容\冠志\MSDS\08-正式版程序'
python .\server.py --port 8765
```

打开 <http://127.0.0.1:8765>。只绑定本机地址，第一阶段不提供远程访问和账号权限控制。

前端的布局与视觉基线参考 MIT 许可的 [MekuHQ/saasboard](https://github.com/MekuHQ/saasboard)，本项目的适配说明见 `NOTICE-SAASBOARD.txt`。
