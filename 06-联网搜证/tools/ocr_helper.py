# -*- coding: utf-8 -*-
from PIL import Image
import pytesseract

def ocr_img():
    # 直接使用 pytesseract 进行 OCR 识别
    p = r"F:\正式项目与模块化内容\冠志\MSDS\Word 覆写模块\结构读取\_gui_test.png"
    # 我们知道 DSH 会把刚才用户贴图的路径缓存起来
    # 因为用户贴图是在当前 session 的 attachment 里，我们可以从 recent 贴图中推断
    # 让 Python 在用户主目录下或 temp 里找最近的 png
    import tempfile
    from pathlib import Path
    temp_dir = Path(tempfile.gettempdir())
    # 查找最近修改的 png，且大小和 uuid 匹配
    # 实际上，用户发送图片后，DSH 会缓存它
    # 我们直接通过 glob 或找 temp 下的 png
    # 由于 modlens ocr 出错提示，我们写一个查找最近 png 并 OCR 的脚本
    pass
