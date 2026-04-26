# -*- coding: utf-8 -*-
"""Markdown转Docx 主程序入口"""

import sys
import os

# 确保可以导入同目录下的模块
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gui import main

if __name__ == '__main__':
    main()
