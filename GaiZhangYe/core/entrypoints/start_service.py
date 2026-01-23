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
    """清除旧的Python进程和占用指定端口的进程"""
    print("[1/3] 正在检查旧的Python进程...")

    # 要检查的端口
    ports_to_check = [5001]

    try:
        # Windows系统
        if sys.platform == "win32":
            # 只杀死占用端口的进程，不再杀死所有Python进程，避免杀死自身
            print("✓ 跳过杀死所有Python进程，改为只检查占用端口的进程")

            # 再检查并杀死占用指定端口的进程
            for port in ports_to_check:
                try:
                    # 查找占用端口的进程ID
                    result = subprocess.run(
                        ["netstat", "-ano", "-p", "tcp"],
                        capture_output=True, text=True
                    )

                    # 分析结果找到对应的PID
                    for line in result.stdout.split('\n'):
                        if f":{port} " in line and "LISTENING" in line:
                            parts = line.split()
                            pid = parts[-1]
                            print(f"  发现占用端口 {port} 的进程PID: {pid}")

                            # 杀死该进程
                            subprocess.run(["taskkill", "/F", "/PID", pid], capture_output=True)
                            print(f"  ✓ 已杀死占用端口 {port} 的进程")
                except Exception as e:
                    print(f"  × 检查端口 {port} 失败: {e}")

        # Linux/macOS系统
        elif sys.platform in ["linux", "darwin"]:
            # 检查并杀死占用指定端口的进程
            for port in ports_to_check:
                try:
                    # 查找占用端口的进程ID
                    result = subprocess.run(
                        ["lsof", "-i", f":{port}"],
                        capture_output=True, text=True
                    )

                    # 分析结果找到对应的PID
                    for line in result.stdout.split('\n')[1:]:  # 跳过表头
                        if line:
                            parts = line.split()
                            pid = parts[1]
                            print(f"  发现占用端口 {port} 的进程PID: {pid}")

                            # 杀死该进程
                            subprocess.run(["kill", "-9", pid], capture_output=True)
                            print(f"  ✓ 已杀死占用端口 {port} 的进程")
                except Exception as e:
                    print(f"  × 检查端口 {port} 失败: {e}")

    except Exception as e:
        print("× 清除旧进程失败:", str(e))

def start_service():
    """启动服务"""
    print("[2/3] 正在启动HTML服务...")

    # 将项目根目录添加到Python路径
    if getattr(sys, 'frozen', False):
        # PyInstaller打包后的情况
        project_root = sys._MEIPASS
    else:
        # 开发模式的情况
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
            time.sleep(1)
            webbrowser.open('http://localhost:5001')

        import threading
        browser_thread = threading.Thread(target=open_browser)
        browser_thread.daemon = True
        browser_thread.start()

        # 启动服务
        app.run(
            host='0.0.0.0',
            port=5001,
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

    # 尝试杀死旧进程
    kill_old_processes()

    start_service()

if __name__ == "__main__":
    main()
