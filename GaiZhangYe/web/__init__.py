#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Web package initializer: create_app factory and blueprint registration."""
import os
import sys
import uuid
from flask import Flask
from flask_restx import Api


def create_app():
    # 处理PyInstaller打包后的路径问题
    if getattr(sys, 'frozen', False):
        app_root = sys._MEIPASS
    else:
        app_root = os.path.abspath(os.path.dirname(__file__))

    app = Flask(
        __name__,
        template_folder=os.path.join(app_root, "templates"),
        static_folder=os.path.join(app_root, "static"),
    )

    # 全局会话ID
    app.config['APP_SESSION_ID'] = str(uuid.uuid4())

    # 配置Flask-RESTX
    api = Api(
        app,
        title='盖章页工具API',
        version='1.0',
        description='盖章页处理工具的API接口文档',
        doc='/api/doc',  # API文档访问路径
        prefix='/api',  # API前缀
    )

    # 注册蓝图和API命名空间
    from .routes.pages import pages_bp
    from .routes.api import api_ns

    # 注册页面蓝图
    app.register_blueprint(pages_bp)
    # 将API命名空间添加到Flask-RESTX的Api对象中
    api.add_namespace(api_ns)

    # 调试路由信息
    @app.route('/debug/routes')
    def list_routes():
        output = []
        for rule in app.url_map.iter_rules():
            output.append(f"{rule.endpoint}: {rule.rule}")
        return "\n".join(sorted(output)), 200, {'Content-Type': 'text/plain'}

    return app
