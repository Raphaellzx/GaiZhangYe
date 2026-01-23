#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Web package initializer: create_app factory and blueprint registration."""
import os
import sys
import uuid
from flask import Flask


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

    # 注册蓝图
    from .routes.pages import pages_bp
    from .routes.api import api_bp

    app.register_blueprint(pages_bp)
    app.register_blueprint(api_bp)

    return app
