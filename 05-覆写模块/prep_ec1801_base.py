# -*- coding: utf-8 -*-
"""预处理 PEA-4139 模板 → EC-1801 标准格式底版（去残留版）：
删除模板 PEA-4139 产品特有的残留行（内容污染源）：
  - S8: 氟化橡胶/丁基橡胶/丁腈橡胶 手套材质行
  - S11: "该产品无可用的毒理学研究。" + "以下为...参考数据" 引导行
  - S12: "该产品无可用的生态毒理学研究。" + "以下为...参考数据" 引导行
  - S13: "必须遵守适用的国标..." + "在欧盟领域内废弃..." 引导行
  - S12: 第二个重复的 12.1 生态毒性行
同时清理 S8 手部防护 label 的制表符残留。
"""
import sys
sys.path.insert(0, r"F:\正式项目与模块化内容\冠志\MSDS\02-检索系统")
import docx
from msds_overwrite_engine import section_tables, qn

SRC = r"F:\正式项目与模块化内容\冠志\MSDS\03-数据库\正式库\推导方案\PEA-4139 MSDS_CN 冠志 模板.docx"
DST = r"F:\正式项目与模块化内容\冠志\MSDS\03-数据库\测试库\输出库\EC-1801 底版_冠志.docx"

def set_cell_text(cell_elm, text):
    paras = cell_elm.findall(qn('w:p'))
    if not paras:
        return
    p0 = paras[0]
    rPr = None
    first_r = p0.find(qn('w:r'))
    if first_r is not None:
        rPr = first_r.find(qn('w:rPr'))
    for r in list(p0.findall(qn('w:r'))):
        p0.remove(r)
    for child in list(p0):
        if child.tag != qn('w:pPr'):
            p0.remove(child)
    new_r = p0.makeelement(qn('w:r'), {})
    if rPr is not None:
        new_r.append(rPr)
    new_t = p0.makeelement(qn('w:t'), {})
    new_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    new_t.text = text
    new_r.append(new_t)
    p0.append(new_r)
    for p in paras[1:]:
        cell_elm.remove(p)

def delete_row(tbl, row_idx):
    tr = tbl.rows[row_idx]._tr
    tr.getparent().remove(tr)

def delete_rows_by_text(tbl, match):
    """删除标签格文本匹配 match 的行（倒序删除）。"""
    for ri in range(len(tbl.rows)-1, 0, -1):
        txt = tbl.rows[ri].cells[0].text.strip()
        if match(txt):
            delete_row(tbl, ri)
            print(f'  删 R{ri}: {txt[:50]}')

doc = docx.Document(SRC)
tables = section_tables(doc)

# ---- S8: 清手部防护制表符 + 删手套材质行 ----
t8 = tables[8]
c = t8.rows[3].cells[0]
assert "手部防护" in c.text
set_cell_text(c._tc, "手部防护：")
print('S8 R3 →', repr(c.text))
delete_rows_by_text(t8, lambda t: ("氟化橡胶" in t) or ("丁基橡胶" in t) or ("丁腈橡胶" in t))
print('S8 手套材质行已删，剩余行数:', len(t8.rows))

# ---- S11: 删引导行 ----
t11 = tables[11]
delete_rows_by_text(t11, lambda t: "无可用的毒理学研究" in t or "参考数据" in t)
print('S11 引导行已删，剩余行数:', len(t11.rows))

# ---- S12: 删引导行 + 重复12.1 ----
t12 = tables[12]
delete_rows_by_text(t12, lambda t: "无可用的生态毒理学研究" in t or "参考数据" in t)
# 删除第二个 12.1 生态毒性
seen = False
for ri in range(1, len(t12.rows)):
    txt = t12.rows[ri].cells[0].text.strip()
    if txt.startswith("12.1"):
        if seen:
            delete_row(t12, ri)
            print(f'S12 R{ri} 删重复12.1: {txt!r}')
            break
        seen = True
print('S12 剩余行数:', len(t12.rows))

# ---- S13: 删欧盟引导行 ----
t13 = tables[13]
delete_rows_by_text(t13, lambda t: "必须遵守适用的国标" in t or "欧盟领域内废弃" in t or "EWC" in t)
print('S13 引导行已删，剩余行数:', len(t13.rows))

doc.save(DST)
print('已保存底版:', DST)
