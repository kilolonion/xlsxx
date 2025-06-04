# -*- coding: utf-8 -*-
"""
XLS批量转换工具 - Streamlit应用主入口
"""

import sys
import os

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 导入主界面
from ui.main_interface import main

if __name__ == "__main__":
    main() 