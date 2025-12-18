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
    # 直接列出func1下的固定目录，确保目录按钮总是显示
    folders = ["Nostamped_Word", "Nostamped_PDF", "Stamped_Pages", "Temp"]

    # 获取Nostamped_Word目录中的Word文件
    word_files = []
    nostamped_word_dir = file_manager.get_func1_dir("nostamped_word")
    if nostamped_word_dir.exists():
        import os
        print(f"目录存在: {nostamped_word_dir}")
        file_count = 0
        for filename in os.listdir(nostamped_word_dir):
            print(f"检查文件: {filename}")
            if filename.endswith(('.docx', '.doc')):
                print(f"添加文件: {filename}")
                word_files.append({
                    'name': filename,
                    'stem': os.path.splitext(filename)[0]
                })
                file_count += 1
        print(f"总计找到: {file_count}个文件")

    return render_template('pages/prepare_stamp.html', folders=folders, word_files=word_files)

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

@app.route('/api/open-directory')
def open_directory():
    """打开指定目录"""
    import sys
    dir_name = request.args.get('dir_name')
    try:
        if dir_name:
            # 获取目录路径 - 支持func1和func2
            # 目录名映射
            func1_map = {
                "Nostamped_Word": "nostamped_word",
                "Nostamped_PDF": "nostamped_pdf",
                "Stamped_Pages": "stamped_pages",
                "Temp": "temp"
            }
            func2_map = {
                "Images": "images",
                "TargetFiles": "target_files",
                "Result_Word": "result_word",
                "Result_PDF": "result_pdf"
            }

            # 检查func1目录
            if dir_name in func1_map:
                dir_path = file_manager.get_func1_dir(func1_map[dir_name])
            # 检查func2目录
            elif dir_name in func2_map:
                dir_path = file_manager.get_func2_dir(func2_map[dir_name])
            else:
                return jsonify({"success": False, "error": f"无效的目录名: {dir_name}"})

            if dir_path.exists():
                dir_str = str(dir_path)

                # 使用系统默认方式打开目录
                if sys.platform == "win32":
                    os.startfile(dir_str)
                elif sys.platform == "darwin":
                    os.system(f"open '{dir_str}'")
                else:  # Linux
                    os.system(f"xdg-open '{dir_str}'")

                return jsonify({"success": True, "message": "目录已打开"})
            else:
                return jsonify({"success": False, "error": f"目录不存在: {dir_name}"})
        else:
            return jsonify({"success": False, "error": "未提供目录名"})
    except Exception as e:
        logger.error(f"打开目录失败: {str(e)}")
        return jsonify({"success": False, "error": f"打开目录失败: {str(e)}"})

@app.route('/api/prepare-stamp', methods=['POST'])
def prepare_stamp():
    """准备盖章页"""
    try:
        from GaiZhangYe.core.stamp_prepare import StampPrepareService

        target_pages = request.get_json()

        if not target_pages:
            return jsonify({"success": False, "error": "没有提供页面范围"})

        logger.info(f"准备处理页面范围: {target_pages}")

        # 创建服务实例并执行
        stamp_service = StampPrepareService()
        result_files = stamp_service.run(target_pages)

        # 转换为字符串路径
        result_files_str = [str(file) for file in result_files]

        return jsonify({
            "success": True,
            "message": "盖章页准备完成",
            "files": result_files_str
        })
    except Exception as e:
        logger.error(f"准备盖章页失败: {str(e)}", exc_info=True)
        return jsonify({"success": False, "error": f"准备盖章页失败: {str(e)}"})

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


# --------------------------
# 新增: 数据沟通API
# --------------------------
@app.route('/api/refresh-data', methods=['POST'])
def api_refresh_data():
    """刷新数据文件接口"""
    try:
        from GaiZhangYe.core.data_communication import get_data_service
        data_service = get_data_service()
        if data_service.scan_business_data():
            return jsonify({
                'success': True,
                'message': '数据文件已重新生成'
            })
        return jsonify({
            'success': False,
            'error': '数据文件重新生成失败'
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        })


@app.route('/api/func1/data', methods=['GET'])
def api_get_func1_data():
    """获取func1的target_pages数据"""
    try:
        from GaiZhangYe.core.data_communication import get_data_service
        data = get_data_service().get_func1_data()
        return jsonify({
            'success': True,
            'data': data
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        })


@app.route('/api/func1/data', methods=['POST'])
def api_save_func1_data():
    """保存func1的target_pages数据"""
    try:
        from GaiZhangYe.core.data_communication import get_data_service
        data = request.get_json()
        if get_data_service().save_func1_data(data):
            return jsonify({
                'success': True
            })
        return jsonify({
            'success': False,
            'error': '保存失败'
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        })


@app.route('/api/func2/data', methods=['GET'])
def api_get_func2_data():
    """获取func2的stamp_config数据"""
    try:
        from GaiZhangYe.core.data_communication import get_data_service
        data = get_data_service().get_func2_data()
        return jsonify({
            'success': True,
            'data': data
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        })


@app.route('/api/func2/data', methods=['POST'])
def api_save_func2_data():
    """保存func2的stamp_config数据"""
    try:
        from GaiZhangYe.core.data_communication import get_data_service
        data = request.get_json()
        if get_data_service().save_func2_data(data):
            return jsonify({
                'success': True
            })
        return jsonify({
            'success': False,
            'error': '保存失败'
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        })


if __name__ == '__main__':
    logger.info("启动HTML可视化服务...")
    logger.info("服务将在 http://localhost:5001 启动")
    logger.info("按 Ctrl+C 停止服务")

    app.run(
        host='0.0.0.0',
        port=5001,
        debug=False
    )

