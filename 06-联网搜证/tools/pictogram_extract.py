# -*- coding: utf-8 -*-
"""从 GB 30000 系列标准提取全部 GHS 象形图.
输出: 标准原文归档/象形图库/原始提取/{标准号}_p{页码}_{序号}.png
      标准原文归档/象形图库/pictograms_index.json  (每张图的来源+上下文)
方法: fitz get_pixmap(xref) 提取位图, 页面 bbox 记录位置. 用 xref 去重."""
import fitz, glob, os, json, re
from pathlib import Path

SRC = Path(r'F:/正式项目与模块化内容/冠志/MSDS/04-推断引擎/法规匹配库/标准原文归档/GB30000系列')
OUT = Path(r'F:/正式项目与模块化内容/冠志/MSDS/04-推断引擎/法规匹配库/标准原文归档/象形图库/原始提取')
OUT.mkdir(parents=True, exist_ok=True)

# 标准号 → 危险类别 (GB 30000 系列)
STD_CAT = {
    '2':'爆炸物', '3':'易燃气体', '4':'气溶胶', '5':'氧化性气体', '6':'加压气体',
    '7':'易燃液体', '8':'易燃固体', '9':'自反应物质', '10':'自燃液体', '11':'自燃固体',
    '12':'自热物质', '13':'遇水放出易燃气体', '14':'氧化性液体', '15':'氧化性固体',
    '16':'有机过氧化物', '17':'金属腐蚀物', '18':'急性毒性', '19':'皮肤腐蚀/刺激',
    '20':'严重眼损伤/刺激', '21':'呼吸道/皮肤致敏', '22':'生殖细胞致突变性', '23':'致癌性',
    '24':'生殖毒性', '25':'STOT单次', '26':'STOT反复', '27':'吸入危害', '28':'危害水生环境',
    '29':'危害臭氧层', '30':'退敏爆炸物', '1':'通则', '.1':'通则',
}

def main():
    files = sorted(glob.glob(str(SRC / 'GB*.pdf')))
    index = []
    seen_xrefs = set()
    seen_boxes = []   # 无 xref 时的 bbox 去重
    total = 0
    for f in files:
        std = os.path.basename(f).replace('.pdf', '')
        # 提取标准分部分号
        m = re.search(r'GB30000\.(\d+)', std) or re.search(r'GB 30000\.(\d+)', std)
        part = m.group(1) if m else '1'
        cat = STD_CAT.get(part, '未知')
        doc = fitz.open(f)
        for pno in range(len(doc)):
            page = doc[pno]
            txt = page.get_text()
            # 只处理含象形图/标签要素的页
            if not ('象形图' in txt or '无象形图' in txt or '标签要素' in txt): continue
            # 页面上的图片信息(含 bbox)
            img_infos = page.get_image_info()
            for ii in img_infos:
                xref = ii.get('xref')
                bbox = ii.get('bbox', [0, 0, 0, 0])
                # bbox 去重(无 xref 时退化: bbox 近似相同视为同一图)
                key = (round(bbox[0], 1), round(bbox[1], 1))
                if any(abs(a-key[0]) < 2 and abs(b-key[1]) < 2 for a, b in seen_boxes):
                    continue
                if xref and xref in seen_xrefs:
                    continue
                if xref:
                    seen_xrefs.add(xref)
                seen_boxes.append(key)
                rect = fitz.Rect(bbox)
                try:
                    if xref:
                        pix = fitz.Pixmap(doc, xref)
                        if pix.colorspace and pix.colorspace.n > 3:
                            pix = fitz.Pixmap(fitz.csRGB, pix)
                    else:
                        # 矢量/内嵌图: 按 bbox 区域渲染, 目标短边约 600px
                        zoom = min(600 / max(rect.width, 1), 600 / max(rect.height, 1), 3.0)
                        pix = page.get_pixmap(clip=rect, matrix=fitz.Matrix(zoom, zoom))
                    if pix.width < 20 or pix.height < 20: continue  # 过滤小装饰图
                    fname = f'{std}_p{pno+1}_{total}.png'
                    pix.save(str(OUT / fname))
                    total += 1
                    index.append({
                        'file': fname, 'std': std, 'part': part, 'cat': cat,
                        'page': pno+1, 'xref': xref, 'size': f'{pix.width}x{pix.height}',
                        'bbox': [round(x,1) for x in bbox],
                    })
                    pix = None
                except Exception as e:
                    pass
        doc.close()
    json.dump(index, open(OUT.parent / 'pictograms_index.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(f'提取象形图 {total} 张 → {OUT}')
    from collections import Counter
    c = Counter(i['cat'] for i in index)
    for cat, n in sorted(c.items(), key=lambda x:-x[1]):
        print(f'  {cat}: {n}')

if __name__ == '__main__':
    main()
