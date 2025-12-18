#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动启动脚本 - 解决404问题
"""

import os
import sys
import subprocess
import time
import webbrowser

def kill_old_processes():
    """杀死旧的Python进程"""
    print("[1/3] 正在检查旧的Python进程...")
    try:
        # Windows系统
        if sys.platform == "win32":
            # 尝试使用更简单的方法杀死可能的旧进程
            # 直接杀死所有python.exe进程中的app_api.py，使用taskkill的filter
            # 注意：这种方法可能无法完全匹配，但更可靠
            subprocess.run(["taskkill", "/F", "/FI", "IMAGENAME eq python.exe", "/FI", "WINDOWTITLE eq Python"],
                         capture_output=True)
            print("已尝试清除旧的Python进程")
        # Linux/macOS系统
        elif sys.platform in ["linux", "darwin"]:
            subprocess.run(["pkill", "-f", "app_api.py"], capture_output=True)
            print("✓ 已杀死旧的web服务进程")
    except Exception as e:
        print("× 清除旧进程失败或没有旧进程:", str(e))

def start_service():
    """启动服务"""
    print("[2/3] 正在启动HTML服务...")

    # 将项目根目录添加到Python路径
    current_file = os.path.abspath(__file__)
    project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(current_file))))
    sys.path.insert(0, project_root)

    # 打印调试信息
    print(f"项目根目录: {project_root}")
    print(f"Python路径: {sys.path}")

    try:
        from GaiZhangYe.web.app import app

        print("服务将在 http://localhost:5001 启动")

        # 打开浏览器
        def open_browser():
            time.sleep(2)
            webbrowser.open('http://localhost:5001')

        import threading
        browser_thread = threading.Thread(target=open_browser)
        browser_thread.daemon = True
        browser_thread.start()

        # 启动服务
        app.run(
            host='0.0.0.0',
            port=5000,
            debug=False
        )

    except Exception as e:
        print("启动服务失败:", str(e))
        import traceback
        traceback.print_exc()
        sys.exit(1)

def main():
    """主函数"""
    print("=" * 40)
    print("盖章页工具 HTML 可视化服务")
    print("=" * 40)
    print()

    # 不再尝试杀死旧进程
    print("[1/2] 跳过旧进程检查")
    print()

    start_service()

    print("[3/3] 正在打开浏览器...")
    print()



if __name__ == "__main__":
    main()
