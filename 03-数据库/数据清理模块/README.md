# MSDS 成分批量汇总脚本

## 1. 安装

```bash
python -m venv .venv
```

Windows:

```bash
.venv\Scripts\activate
pip install -r requirements.txt
```

macOS / Linux:

```bash
source .venv/bin/activate
pip install -r requirements.txt
```

## 2. 运行

```bash
python msds_extract.py "D:\MSDS文件夹" -o "D:\MSDS成分汇总.xlsx"
```

macOS / Linux 示例：

```bash
python msds_extract.py "/Users/name/MSDS" -o "/Users/name/MSDS成分汇总.xlsx"
```

## 3. 输出

- `MSDS汇总`：产品、成份、标准、含量及来源信息
- `异常与待处理`：扫描 PDF、未找到第 3 节、解析失败等

## 4. 已知限制

- 扫描版 PDF 不含文字层，需要先 OCR。
- 老式 `.doc`、`.xls` 建议先转成 `.docx`、`.xlsx`。
- 供应商模板差异较大，应根据样本文档调整脚本中的表头别名。
- “标准”默认按 CAS/EC/标准号/标识号理解。
