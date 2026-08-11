# -*- coding: utf-8 -*-
"""启动 GUI 并打开 10 份中文冠志测试库 (原正式库不动)."""
from __future__ import annotations
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from gui.main_window import MainWindow

TEST_DB = Path(r"F:\正式项目与模块化内容\Word 覆写模块\数据库\测试库_10份冠志\正式筛选库.db")


class TestApp(MainWindow):
    def __init__(self):
        super().__init__()
        self.screened.open_db(str(TEST_DB))
        self.notebook.select(2)


if __name__ == "__main__":
    TestApp().mainloop()
