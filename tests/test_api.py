#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
直接在Flask应用上下文中测试API的脚本
"""

import sys
import os
from pathlib import Path

# 添加项目路径到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

def test_api_in_context():
    try:
        from GaiZhangYe.web import create_app
        app = create_app()

        print("OK Flask应用创建成功")

        # 在Flask应用上下文中测试API
        with app.test_request_context('/api/status'):
            try:
                from GaiZhangYe.web.routes.api import Status
                print("OK Status资源类导入成功")

                # 测试API端点
                print("\n=== 测试API状态接口 ===")
                client = app.test_client()
                response = client.get('/api/status')
                print(f"状态码: {response.status_code}")
                print(f"响应内容: {response.data.decode('utf-8')}")

                print("\n=== 测试API会话ID接口 ===")
                response = client.get('/api/session-id')
                print(f"状态码: {response.status_code}")
                print(f"响应内容: {response.data.decode('utf-8')}")

                print("\n=== 测试API目录接口 ===")
                response = client.get('/api/directories')
                print(f"状态码: {response.status_code}")
                print(f"响应内容: {response.data.decode('utf-8')}")

            except Exception as e:
                print(f"ERROR 在Flask应用上下文中测试API失败: {e}")
                import traceback
                print(traceback.format_exc())

    except Exception as e:
        print(f"ERROR 创建Flask应用失败: {e}")
        import traceback
        print(traceback.format_exc())

def test_api_routes():
    try:
        from GaiZhangYe.web import create_app
        app = create_app()

        print("\n=== 打印所有注册的Flask路由 ===")
        for rule in app.url_map.iter_rules():
            print(f"{rule.endpoint}: {rule.rule}")

    except Exception as e:
        print(f"ERROR 获取Flask路由失败: {e}")
        import traceback
        print(traceback.format_exc())

if __name__ == "__main__":
    print("测试Flask应用API")
    print("=" * 50)

    test_api_routes()
    test_api_in_context()