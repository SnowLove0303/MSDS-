# 格式签名验证 + 残留扫描

覆写完成后做三重验证：闭环校验（引擎自动）→ 残留扫描 → 格式签名对比。

## 1. 残留扫描

输出不得含模板产品特有残留关键词。用结构读取读回后全文扫描：

```python
import docx, re
out = docx.Document("<输出.docx>")
F = "\n".join(c.text for tb in out.tables for row in tb.rows for c in row.cells)
bad = ["二乙二醇单丁醚", "乳白色液体", "P210", "H316", "非危险品",
       "EN 374-3", "氟化橡胶", "测试", "成分1", "成分2", ">n"]
for kw in bad:
    if re.search(re.escape(kw), F):
        print(f"残留 [{kw}]")
# 注意：S3 成分 CAS="商业机密" 可能是源真实数据，不要列入残留关键词
```

## 2. 标签格格式签名（双 run 结构）

模板标准：**序号 run（Times New Roman 10.5pt 加粗）+ 中文 run（宋体/Arial 12pt 加粗）**。

```python
import docx, re
from docx.oxml.ns import qn

def cell_runs(cell):
    """标签格首段 run 列表 [(text, font, size_pt, bold)]"""
    out = []
    for p in cell.paragraphs:
        for r in p._p.iterchildren():
            if r.tag == qn("w:r"):
                rpr = r.find(qn("w:rPr"))
                fname = sz = bold = None
                if rpr is not None:
                    rf = rpr.find(qn("w:rFonts"))
                    if rf is not None:
                        fname = rf.get(qn("w:ascii")) or rf.get(qn("w:hAnsi"))
                    s = rpr.find(qn("w:sz"))
                    if s is not None:
                        sz = int(s.get(qn("w:val"))) / 2
                    bold = rpr.find(qn("w:b")) is not None
                t = "".join(x.text or "" for x in r.findall(qn("w:t")))
                if t:
                    out.append((t, fname, sz, bold))
        break
    return out

out = docx.Document("<输出.docx>")
for tb in out.tables:
    for ri, row in enumerate(tb.rows):
        if ri == 0 or len(row._tr.findall(qn("w:tc"))) < 2:
            continue  # 跳过标题行与 note 单格行
        runs = cell_runs(row.cells[0])
        text = "".join(x[0] for x in runs)
        m = re.match(r"^(\d+(?:\.\d+)*)\s*(.*)", text, re.S)
        if not (m and m.group(2).strip()) or len(runs) < 2:
            continue
        # 中文 run = 最后一个含非空白文本的 run（跳过尾部对齐空格 run）
        cn = next((x for x in reversed(runs) if x[0].strip()), None)
        if not (runs[0][2] == 10.5 and cn and cn[2] == 12.0 and cn[3] is True):
            print(f"格式异常 T{ri}: {text[:20]!r} runs={[(x[0][:8], x[2]) for x in runs]}")
```

> 判定要点：跳过纯尾部空格 run 找中文 run；S9 的 9.2/9.6/9.7 是
> `[序号 TNR10.5][中文 Arial12][尾空格 TNR10.5]` 三 run，中文仍须 12pt。

## 3. 与底版逐行对比（模板驱动一致性）

按字段标签匹配底版与输出，验证复用行标签格 100% 继承：

```python
import docx, re
from docx.oxml.ns import qn

def sig(tc):
    out = []
    for p in tc.findall(qn("w:p")):
        for r in p.iterchildren():
            if r.tag == qn("w:r"):
                rpr = r.find(qn("w:rPr"))
                f = sz = b = None
                if rpr is not None:
                    rf = rpr.find(qn("w:rFonts"))
                    if rf is not None: f = rf.get(qn("w:ascii")) or rf.get(qn("w:hAnsi"))
                    s = rpr.find(qn("w:sz"))
                    if s is not None: sz = int(s.get(qn("w:val")))/2
                    b = rpr.find(qn("w:b")) is not None
                t = "".join(x.text or "" for x in r.findall(qn("w:t")))
                if t: out.append((t, f, sz, b))
        break
    return out

def strip_seq(txt):
    m = re.match(r"^(\d+(?:\.\d+)*\s*)(.*)$", txt, re.S)
    return (m.group(1), m.group(2)) if m else ("", txt)

def norm(s): return re.sub(r"\s+", "", str(s or ""))

tpl = docx.Document("<底版.docx>")
out = docx.Document("<输出.docx>")
label_diff = fmt_diff = total = 0
for tb_t, tb_o in zip(tpl.tables, out.tables):
    t_map = {}
    for r in tb_t.rows[1:]:
        tcs = r._tr.findall(qn("w:tc"))
        if len(tcs) < 2: continue
        s = sig(tcs[0]); txt = "".join(x[0] for x in s)
        _seq, label = strip_seq(txt)
        if norm(label): t_map.setdefault(norm(label), s)
    for r in tb_o.rows[1:]:
        tcs = r._tr.findall(qn("w:tc"))
        if len(tcs) < 2: continue
        s = sig(tcs[0]); txt = "".join(x[0] for x in s)
        _seq, label = strip_seq(txt)
        if not norm(label): continue
        total += 1
        ts = t_map.get(norm(label))
        if ts is None:
            print(f"[新增] {label[:18]}")            # 模板没有、写入项新增，正常
            continue
        if "".join(x[0] for x in ts) != txt:
            label_diff += 1; print(f"[标签差异] {label[:18]}")
        if [(x[1], x[2], x[3]) for x in ts] != [(x[1], x[2], x[3]) for x in s]:
            fmt_diff += 1; print(f"[格式差异] {label[:18]}")
print(f"对比 {total} 行 | 标签差异 {label_diff} | 格式差异 {fmt_diff}")
```

**通过标准**：标签差异 0、格式差异 0；"新增"只出现在模板本就没有的字段
（如 EC-1801 的 S3 第三成分三亚乙基四胺）。

## 4. 结构读取回读（16 节完整性）

```bash
cd "F:\正式项目与模块化内容\冠志\MSDS\02-检索系统"
python main.py --extract "<输出.docx>" | grep -E "^\[[0-9]+\]"
```

应列出 `[0]`~`[16]` 全部 17 节；S3 成分数正确；S12 生态数据完整。
