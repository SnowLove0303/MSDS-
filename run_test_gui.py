# -*- coding: utf-8 -*-
"""启动 GUI (结构读取 · 检索系统)."""
from __future__ import annotations
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from gui.main_window import MainWindow


if __name__ == "__main__":
    MainWindow().mainloop()
