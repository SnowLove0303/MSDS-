# -*- coding: utf-8 -*-
"""深度智能物理归集：重整标准原文归档与08未分类待核验等所有子目录."""
import os, shutil, re
from pathlib import Path

BASE = Path(r"F:\正式项目与模块化内容\冠志\MSDS\04-推断引擎\法规匹配库")

def run():
    # 规则 1：将 08_未分类_待核验 中的文件根据关键字和模式移动到各 01~07 正式目录
    src_08 = BASE / "08_未分类_待核验"
    if src_08.exists():
        for file in src_08.iterdir():
            if file.is_file():
                name = file.name
                target_dir = None
                
                # 判定规则
                if "GB 30000" in name:
                    target_dir = BASE / "03_技术标准_GHS分类"
                elif "JT/T" in name or "GB 6944" in name or "GB 12268" in name or "TDG" in name:
                    target_dir = BASE / "05_技术标准_运输与标签"
                elif "GB 15603" in name or "GB 12158" in name or "GB 50016" in name or "GB 50140" in name or "GB 17914" in name or "GB 17915" in name or "GB 17916" in name or "储存" in name:
                    target_dir = BASE / "04_技术标准_储存与消防"
                elif "GBZ" in name or "WS " in name or "防毒面具" in name or "呼吸" in name or "口罩" in name or "接触限值" in name:
                    target_dir = BASE / "06_技术标准_职业卫生与急救"
                elif "GB 18597" in name or "GB 18484" in name or "GB 50483" in name or "废物" in name:
                    target_dir = BASE / "07_技术标准_环保与处置"
                elif "GB 30981" in name or "限量" in name or "VOC" in name:
                    target_dir = BASE / "07_技术标准_应用限制"
                elif "GB/T 16483" in name or "GB/T 17519" in name or "GB 15258" in name or "安全标签" in name:
                    target_dir = BASE / "03_技术标准_GHS分类"
                elif "GB/T 21844" in name or "GB/T 21853" in name or "GB/T 21845" in name or "GB/T 21578" in name or "GB/T 21807" in name or "GB/T 21809" in name or "GB/T 21805" in name or "试验" in name or "测定" in name:
                    target_dir = BASE / "06_技术标准_职业卫生与急救"
                
                if target_dir:
                    target_dir.mkdir(parents=True, exist_ok=True)
                    shutil.move(str(file), str(target_dir / name))
                    print(f"整理: 08\\{name} -> {target_dir.name}")

    # 规则 2：将 标准原文归档 中的文件也合并到 03~07 标准目录
    std_archive = BASE / "标准原文归档"
    if std_archive.exists():
        for root, dirs, files in os.walk(std_archive):
            for file in files:
                src_file = Path(root) / file
                name = file
                target_dir = None
                
                if "GB 30000" in name:
                    target_dir = BASE / "03_技术标准_GHS分类"
                elif "JT/T" in name or "GB 6944" in name or "GB 12268" in name or "TDG" in name or "运输" in root:
                    target_dir = BASE / "05_技术标准_运输与标签"
                elif "GB 15603" in name or "GB 12158" in name or "GB 50016" in name or "GB 50140" in name or "GB 17914" in name or "GB 17915" in name or "GB 17916" in name or "储存" in name or "储存" in root or "消防" in root:
                    target_dir = BASE / "04_技术标准_储存与消防"
                elif "GBZ" in name or "WS " in name or "面具" in name or "呼吸" in name or "职业卫生" in root:
                    target_dir = BASE / "06_技术标准_职业卫生与急救"
                elif "GB 18597" in name or "GB 18484" in name or "GB 50483" in name or "环保" in root or "危废" in root:
                    target_dir = BASE / "07_技术标准_环保与处置"
                elif "GB 30981" in name or "限量" in name or "VOC" in name or "应用" in root:
                    target_dir = BASE / "07_技术标准_应用限制"
                elif "GB/T 16483" in name or "GB/T 17519" in name or "GB 15258" in name or "SDS" in root or "框架" in root:
                    target_dir = BASE / "03_技术标准_GHS分类"
                elif "GB/T 21844" in name or "GB/T 21853" in name or "GB/T 21845" in name or "GB/T 21578" in name or "GB/T 21807" in name or "GB/T 21809" in name or "GB/T 21805" in name or "测试" in root:
                    target_dir = BASE / "06_技术标准_职业卫生与急救"
                
                if target_dir:
                    target_dir.mkdir(parents=True, exist_ok=True)
                    dest_path = target_dir / name
                    if not dest_path.exists():
                        shutil.move(str(src_file), str(dest_path))
                        print(f"整理标准: {src_file.relative_to(BASE)} -> {target_dir.name}")
                    else:
                        os.remove(str(src_file))

        # 清空标准原文归档目录
        shutil.rmtree(std_archive)
        print("已清理'标准原文归档'目录。")

if __name__ == "__main__":
    run()
