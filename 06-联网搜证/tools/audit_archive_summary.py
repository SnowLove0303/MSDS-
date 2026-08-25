# -*- coding: utf-8 -*-
"""更新国内外法规归档清单与标准原文归档清单，汇总最新全景状态."""
from pathlib import Path

ROOT = Path(r"F:\正式项目与模块化内容\冠志\MSDS\04-推断引擎\法规匹配库")
LAW_DIR = ROOT / "法规原文归档"
STD_DIR = ROOT / "标准原文归档"

def audit_archive():
    law_files = list(LAW_DIR.rglob("*.docx")) + list(LAW_DIR.rglob("*.pdf")) + list(LAW_DIR.rglob("*.doc"))
    std_files = list(STD_DIR.rglob("*.docx")) + list(STD_DIR.rglob("*.pdf")) + list(STD_DIR.rglob("*.doc"))
    
    print(f"法规文件总数 (docx/pdf/doc): {len(law_files)}")
    print(f"标准文件总数 (docx/pdf/doc): {len(std_files)}")
    print(f"全库总计归档规范文件: {len(law_files) + len(std_files)} 份")

if __name__ == '__main__':
    audit_archive()
