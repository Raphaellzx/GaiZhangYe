#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
盖章页工具 HTML 版本启动脚本
"""

import sys
import os

# 确保能够导入GaiZhangYe模块
sys.path.insert(0, str(os.path.dirname(os.path.abspath(__file__))))

try:
    from GaiZhangYe.web.app import app
except ImportError as e:
    print(f"错误: 导入失败 - {str(e)}")
    print("\n请先安装依赖:")
    print("pip install flask")
    print("\n或安装所有依赖:")
    print("pip install -r requirements.txt")
    sys.exit(1)

if __name__ == "__main__":
    print("盖章页工具 HTML 可视化服务")
    print("=" * 40)
    print("服务将在 http://localhost:5000 启动")
    print("按 Ctrl+C 停止服务")
    print("=" * 40)
    print()

    try:
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
