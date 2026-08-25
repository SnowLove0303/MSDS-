# 完整 CLI 参考

## 覆写引擎（msds_overwrite_engine.py）

```bash
cd "F:\正式项目与模块化内容\冠志\MSDS\05-覆写模块"

python msds_overwrite_engine.py \
  --template "..\数据库\测试库\PEA-4139 MSDS_CN 冠志 模板.docx" \
  --write-items write_items_os1330.json \
  --out "..\数据库\测试库\输出库\OS-1330 标准化输出.docx" \
  --sections 1,2,3,9          # 可选：只处理指定节
```

### CLI 参数

| 参数 | 说明 | 默认 |
|---|---|---|
| `--template` | 模板/底版 docx（必填） | — |
| `--write-items` | 写入项 JSON（必填） | — |
| `--out` | 输出 docx 路径（必填） | — |
| `--sections` | 仅处理指定节，如 `1,3,9` | 全部 |
| `--ref-pick` | 新增行参考行策略：`auto`/`last`/`label:标签`/`row:行号` | `auto` |
| `--keep-structure` | 模板驱动节，如 `2`；不传=全部节模板驱动 | `all` |
| `--empty-policy` | 空值策略：`warn`/`overwrite`/`preserve` | `warn` |
| `--field-map` | 字段映射配置 json | 同目录 `field_maps.json` |

**退出码**：闭环校验失败 → `sys.exit(1)`，日志带 `[FAIL]`。

## 结构读取（结构读取/main.py）

```bash
cd "F:\正式项目与模块化内容\冠志\MSDS\02-检索系统"

python main.py                                        # GUI（左目录 | 右表格+检索）
python main.py --cli "<文件.docx>"                    # 全 16 节摘要 + 告警（anomalies）
python main.py --extract "<文件.docx>"                # 三级父子级树
python main.py --extract "<文件.docx>" --sections 1,3,9
python main.py --extract "<文件.docx>" --query 供应商 --scope label
python main.py --extract "<文件.docx>" --json --out s12.json
python main.py --extract "<文件.docx>" --tsv --out 提取表.tsv    # TSV 带 BOM 可直接 Excel
python main.py --extract "<文件.docx>" "<文件2.docx>" --sections 1,3,9   # 批量
```

`--extract` 参数：`--query 词`、`--scope label|value|all|section`、`--json`/`--tsv`/`--flat`、`--out 文件`、`--sections 1,3,9`。

## 写入项生成

```bash
cd "F:\正式项目与模块化内容\冠志\MSDS\05-覆写模块"

# 单文件：真实 MSDS → 写入项 JSON
python make_real_write_items.py "<真实MSDS.docx>" --out write_items.json

# 批量：目录下所有 docx → 输出目录下同名 JSON
python make_real_write_items.py "<目录>" --out "<输出目录>"
```

生成规则要点：
- S0 只保留可变字段，排除固定标题（物料安全数据表）与自动页码；
- 表格结构节（S1/S9）保留空值字段（模板有对应行则覆写为空）；
- 单格文本节（S2）只保留有值字段（空值骨架不参与覆写）；
- S3 输出 `{"产品类型": "...", "components": [{name, cas, conc}, ...]}`；
- 中文名称补型号：`中文名称 + 空格 + 产品型号`（模型如 `EC-1801`）。

## 底版预处理（prep_ec1801_base.py）

```bash
cd "F:\正式项目与模块化内容\冠志\MSDS\05-覆写模块"
python prep_ec1801_base.py
```

针对当前产品的残留行清理（脚本内写死 SRC/DST 路径），产物放 `数据库\测试库\输出库\`。
改产品时复制脚本并调整匹配关键词（S8 手套材质行 / S11 S12 引导行 / S13 欧盟句 / 重复节行）。
