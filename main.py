#!/usr/bin/env python3
"""
财务数据处理工具 - 主入口

该工具用于自动化处理教育机构收入明细归类，根据sheet1的组织信息
智能匹配sheet2中的明细记录。

使用方法:
    python main.py          # 启动GUI界面
"""

import sys
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from gui.main_window import run_gui


def main():
    """主函数"""
    # 启动GUI界面
    run_gui()


if __name__ == '__main__':
    main()
