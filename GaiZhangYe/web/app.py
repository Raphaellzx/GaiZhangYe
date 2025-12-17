#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HTML可视化服务主入口
"""

import os
import sys
import json
import logging
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_from_directory

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# 添加项目根目录到Python路径
sys.path.insert(0, str(Path(__file__).parent.parent.parent))
sys.path.insert(0, str(Path(__file__).parent.parent))

# 导入核心服务
try:
    from GaiZhangYe.core.stamp_overlay import StampOverlayService
    from GaiZhangYe.core.basic.file_manager import FileManager

    logger.info("核心模块导入成功")
    file_manager = FileManager()
    stamp_service = StampOverlayService()

except ImportError as e:
    logger.error(f"核心模块导入失败: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

# 创建Flask应用
app = Flask(
    __name__,
    template_folder=os.path.abspath(os.path.join(os.path.dirname(__file__), 'templates')),
    static_folder=os.path.abspath(os.path.join(os.path.dirname(__file__), 'static'))
)

# 路由设置
@app.route('/')
def index():
    """主页面"""
    return render_template('index.html')

# 功能页面路由
@app.route('/page/prepare-stamp')
def prepare_stamp_page():
    """准备盖章页页面"""
    return render_template('pages/prepare_stamp.html')

@app.route('/page/stamp-overlay')
def stamp_overlay_page():
    """盖章页覆盖页面"""
    return render_template('pages/stamp_overlay.html')

@app.route('/page/word-to-pdf')
def word_to_pdf_page():
    """Word转PDF页面"""
    return render_template('pages/word_to_pdf.html')

@app.route('/api/status')
def status():
    """服务状态"""
    return jsonify({
        "status": "running",
        "message": "盖章页工具HTML服务已启动",
        "version": "1.0.0"
    })

@app.route('/api/directories')
def get_directories():
    """获取目录信息"""
    try:
        # 获取各个目录的绝对路径
        directories = {
            "nostamped_word": str(file_manager.get_func1_dir("nostamped_word")),
            "nostamped_pdf": str(file_manager.get_func1_dir("nostamped_pdf")),
            "stamped_pages": str(file_manager.get_func1_dir("stamped_pages")),
            "images": str(file_manager.get_func2_dir("images")),
            "target_files": str(file_manager.get_func2_dir("target_files")),
            "result_word": str(file_manager.get_func2_dir("result_word")),
            "result_pdf": str(file_manager.get_func2_dir("result_pdf"))
        }
        return jsonify(directories)

    except Exception as e:
        logger.error(f"获取目录信息失败: {e}")
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    logger.info("启动HTML可视化服务...")
    logger.info("服务将在 http://localhost:5000 启动")
    logger.info("按 Ctrl+C 停止服务")

    app.run(
        host='0.0.0.0',
        port=5000,
        debug=False
    )
