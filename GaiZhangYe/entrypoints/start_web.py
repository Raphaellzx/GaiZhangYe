#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HTML可视化服务快捷启动脚本
"""

import sys
import os
import threading
import webbrowser
import time

# 将项目根目录添加到Python路径
# 获取文件所在目录
current_dir = os.path.dirname(os.path.abspath(__file__))
# 获取GaiZhangYe目录
gai_zhang_ye_dir = os.path.dirname(current_dir)
# 获取项目根目录
project_root = os.path.dirname(gai_zhang_ye_dir)
# 添加到Python路径
sys.path.insert(0, project_root)

from GaiZhangYe.web.app import app

def open_browser():
    """等待服务启动后自动打开浏览器"""
    time.sleep(1)  # 等待2秒，让服务先启动
    webbrowser.open('http://localhost:5000')  # 打开浏览器

if __name__ == "__main__":
    print("启动HTML可视化服务...")
    print("服务将在 http://localhost:5000 启动")
    print("正在自动打开浏览器...")

    # 创建线程打开浏览器
    browser_thread = threading.Thread(target=open_browser)
    browser_thread.daemon = True  # 设置为守护线程，服务停止时自动退出
    browser_thread.start()

    try:
        # 启动服务
        app.run(
            host='0.0.0.0',
            port=5000,
            debug=False
        )
    except KeyboardInterrupt:
        print("\n服务已停止")
        sys.exit(0)
    except Exception as e:
        print(f"\n服务启动失败: {str(e)}")
        sys.exit(1)
