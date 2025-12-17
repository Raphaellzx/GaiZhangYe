# GaiZhangYe/entrypoints/gui.py
"""
图形界面启动入口
"""
from pathlib import Path
import sys

# 确保项目根目录在Python路径中
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from GaiZhangYe.utils.config import init_business_dirs


def main():
    """启动GUI"""
    try:
        # 初始化业务目录
        init_business_dirs()
        print("业务目录初始化完成...")

        # 启动主窗口 - 使用重构后的main_window
        from GaiZhangYe.ui.main_window import main as gui_main
        gui_main()

    except Exception as e:
        print(f"启动失败：{str(e)}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
