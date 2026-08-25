# -*- coding: utf-8 -*-
"""从 GBZ 2.1-2019 PDF 提取表1(化学有害因素 OEL) 为结构化 JSON.
用 find_tables + crop(bbox).extract_text() 正确处理跨行单元格.
表1 位于第 9-33 页(index 8-32), 列: 序号 中文名 英文名 CAS号 MAC PC-TWA PC-STEL 效应 备注.
含第2号修改单(乙草胺 359)."""
import pdfplumber, json
from pathlib import Path

PDF = Path(r'F:/正式项目与模块化内容/冠志/MSDS/04-推断引擎/法规匹配库/标准原文归档/待补标准/GBZ 2.1-2019 工作场所有害因素职业接触限值 第1部分：化学有害因素(含修改单).pdf')
OUT = Path(r'F:/正式项目与模块化内容/冠志/MSDS/04-推断引擎/法规匹配库/推断引擎数据/gbz21_oel.json')

def cl(s):
    if s is None: return None
    return str(s).replace('\n', '').strip() or None

def num(s):
    s = cl(s)
    if s is None: return None
    s = s.replace('（', '').replace('）', '')
    try:
        return float(s) if '.' in s else int(s)
    except ValueError:
        return None

def main():
    pdf = pdfplumber.open(PDF)
    rows = []
    seen = set()
    for pi in range(8, 33):
        page = pdf.pages[pi]
        for tb in page.find_tables():
            for row in tb.rows:
                cells = []
                for bbox in row.cells:
                    if bbox is None: cells.append(None); continue
                    t = page.crop(bbox).extract_text() or ''
                    cells.append(t)
                if not cells or not cl(cells[0]): continue
                no_s = cl(cells[0])
                if not no_s or not no_s.isdigit(): continue
                no = int(no_s)
                cn, en, cas = cl(cells[1]), cl(cells[2]), cl(cells[3])
                if not cas: continue
                if no in seen: continue
                seen.add(no)
                rows.append({
                    'no': no, 'cn': cn, 'en': en, 'cas': cas,
                    'MAC': num(cells[4]), 'PC-TWA': num(cells[5]), 'PC-STEL': num(cells[6]),
                    'effect': cl(cells[7]), 'remark': cl(cells[8]),
                })
    # 第2号修改单: 乙草胺
    rows.append({'no': 359, 'cn': '乙草胺', 'en': 'Acetochlor', 'cas': '34256-82-1',
                 'MAC': None, 'PC-TWA': 0.12, 'PC-STEL': None, 'effect': '肝、肾损伤', 'remark': None})
    rows.sort(key=lambda r: r['no'])
    OUT.write_text(json.dumps(rows, ensure_ascii=False, indent=1), encoding='utf-8')
    nos = [r['no'] for r in rows]
    missing = [i for i in range(1, max(nos)+1) if i not in nos]
    print('总条目:', len(rows), ' 序号', min(nos), '~', max(nos))
    print('缺失序号:', missing)
    print('有MAC:', sum(1 for r in rows if r['MAC']), ' PC-TWA:', sum(1 for r in rows if r['PC-TWA']), ' PC-STEL:', sum(1 for r in rows if r['PC-STEL']))
    for r in rows:
        if r['no'] in (3, 10, 41, 111, 224, 359):
            print(' ', r['no'], r['cn'], '|', r['cas'], '| MAC', r['MAC'], 'TWA', r['PC-TWA'], 'STEL', r['PC-STEL'], '| 备注', r['remark'])

if __name__ == '__main__':
    main()
