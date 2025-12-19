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
    template_folder=os.path.abspath(
        os.path.join(os.path.dirname(__file__), "templates")
    ),
    static_folder=os.path.abspath(os.path.join(os.path.dirname(__file__), "static")),
)


# 路由设置
@app.route("/")
def index():
    """主页面"""
    return render_template("index.html")


# 功能页面路由
@app.route("/page/prepare-stamp")
def prepare_stamp_page():
    """准备盖章页页面"""
    # 直接列出func1下的固定目录，确保目录按钮总是显示
    folders = ["Nostamped_Word", "Nostamped_PDF", "Stamped_Pages", "Temp"]

    # 获取Nostamped_Word目录中的Word文件
    word_files = []
    nostamped_word_dir = file_manager.get_func1_dir("nostamped_word")
    if nostamped_word_dir.exists():
        import os
        from GaiZhangYe.core.basic.word_processor import WordProcessor

        word_processor = WordProcessor()

        for filename in os.listdir(nostamped_word_dir):
            if filename.endswith((".docx", ".doc")):
                file_path = nostamped_word_dir / filename

                # 获取文件页数
                page_count = None
                try:
                    page_count = word_processor.get_word_page_count(file_path)
                except Exception as e:
                    print(f"获取文件 {filename} 页数失败: {e}")

                word_files.append(
                    {
                        "name": filename,
                        "stem": os.path.splitext(filename)[0],
                        "page_count": page_count,
                    }
                )

    return render_template(
        "pages/prepare_stamp.html", folders=folders, word_files=word_files
    )


@app.route("/page/stamp-overlay")
def stamp_overlay_page():
    """盖章页覆盖页面"""
    import os
    from pathlib import Path

    # 获取TargetFiles目录中的Word文件信息（包含总页数）
    word_files = []
    target_files_dir = file_manager.get_func2_dir("target_files")

    # 导入WordProcessor获取页数
    from GaiZhangYe.core.basic.word_processor import WordProcessor

    wp = WordProcessor()

    if target_files_dir.exists():
        for filename in os.listdir(target_files_dir):
            if filename.endswith((".docx", ".doc")):
                file_path = Path(target_files_dir) / filename

                # 获取文件总页数
                try:
                    total_pages = wp.get_word_page_count(file_path)
                    # 如果页数获取失败或为0，默认设为1
                    total_pages = total_pages if total_pages > 0 else 1
                except Exception as e:
                    print(f"获取文件页数失败 {filename}: {e}")
                    total_pages = 1

                # 保存文件名和总页数
                word_files.append({"name": filename, "total_pages": total_pages})

    # 获取images目录中的图片文件
    image_files = []
    images_dir = file_manager.get_func2_dir("images")
    if images_dir.exists():
        for filename in os.listdir(images_dir):
            if filename.endswith((".png", ".jpg", ".jpeg", ".bmp")):
                image_files.append(filename)

    return render_template(
        "pages/stamp_overlay.html", word_files=word_files, image_files=image_files
    )


@app.route("/page/word-to-pdf")
def word_to_pdf_page():
    """Word转PDF页面"""
    return render_template("pages/word_to_pdf.html")


@app.route("/api/word-to-pdf", methods=["POST"])
def word_to_pdf():
    """Word转PDF接口"""
    try:
        from GaiZhangYe.core.batch_convert import BatchConvertService

        data = request.get_json()
        input_dir = (
            Path(data.get("input_dir"))
            if data.get("input_dir")
            else file_manager.get_func2_dir("target_files")
        )
        output_dir = (
            Path(data.get("output_dir"))
            if data.get("output_dir")
            else file_manager.get_func2_dir("result_pdf")
        )

        # 创建服务实例并执行
        convert_service = BatchConvertService()
        result_files = convert_service.run(input_dir, output_dir)

        # 转换为字符串路径
        result_files_str = [str(file) for file in result_files]

        return jsonify(
            {
                "success": True,
                "message": f"转换完成！共生成 {len(result_files_str)} 个PDF文件",
                "output_dir": str(output_dir),
                "files": result_files_str,
            }
        )
    except Exception as e:
        logger.error(f"Word转PDF失败: {str(e)}", exc_info=True)
        return jsonify({"success": False, "error": f"转换失败: {str(e)}"})


@app.route("/api/status")
def status():
    """服务状态"""
    return jsonify(
        {
            "status": "running",
            "message": "盖章页工具HTML服务已启动",
        }
    )


@app.route("/api/open-directory")
def open_directory():
    """打开指定目录"""
    import sys

    dir_name = request.args.get("dir_name")
    try:
        if dir_name:
            # 获取目录路径 - 支持func1和func2
            # 目录名映射
            func1_map = {
                "Nostamped_Word": "nostamped_word",
                "Nostamped_PDF": "nostamped_pdf",
                "Stamped_Pages": "stamped_pages",
            }
            func2_map = {
                "Images": "images",
                "TargetFiles": "target_files",
                "Result_Word": "result_word",
                "Result_PDF": "result_pdf",
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


@app.route("/api/prepare-stamp", methods=["POST"])
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

        return jsonify(
            {"success": True, "message": "盖章页准备完成", "files": result_files_str}
        )
    except Exception as e:
        logger.error(f"准备盖章页失败: {str(e)}", exc_info=True)
        return jsonify({"success": False, "error": f"准备盖章页失败: {str(e)}"})


@app.route("/api/directories")
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
            "result_pdf": str(file_manager.get_func2_dir("result_pdf")),
        }
        return jsonify(directories)

    except Exception as e:
        logger.error(f"获取目录信息失败: {e}")
        return jsonify({"error": str(e)}), 500


# --------------------------
# 数据沟通API
# --------------------------
@app.route("/api/refresh-data", methods=["POST"])
def api_refresh_data():
    """刷新数据文件接口"""
    try:
        from GaiZhangYe.core.data_communication import get_data_service

        data_service = get_data_service()
        if data_service.scan_business_data():
            return jsonify({"success": True, "message": "数据文件已重新生成"})
        return jsonify({"success": False, "error": "数据文件重新生成失败"})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route("/api/func1/data", methods=["GET"])
def api_get_func1_data():
    """获取func1的target_pages数据"""
    try:
        from GaiZhangYe.core.data_communication import get_data_service

        data = get_data_service().get_func1_data()
        return jsonify({"success": True, "data": data})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route("/api/func1/data", methods=["POST"])
def api_save_func1_data():
    """保存func1的target_pages数据"""
    try:
        from GaiZhangYe.core.data_communication import get_data_service

        data = request.get_json()
        if get_data_service().save_func1_data(data):
            return jsonify({"success": True})
        return jsonify({"success": False, "error": "保存失败"})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route("/api/func2/data", methods=["GET"])
def api_get_func2_data():
    """获取func2的stamp_config数据"""
    try:
        from GaiZhangYe.core.data_communication import get_data_service

        data = get_data_service().get_func2_data()
        return jsonify({"success": True, "data": data})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route("/api/func2/data", methods=["POST"])
def api_save_func2_data():
    """保存func2的stamp_config数据"""
    try:
        from GaiZhangYe.core.data_communication import get_data_service

        data = request.get_json()
        if get_data_service().save_func2_data(data):
            return jsonify({"success": True})
        return jsonify({"success": False, "error": "保存失败"})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route("/api/extract-images-from-pdf", methods=["POST"])
def extract_images_from_pdf():
    """从PDF中提取图片到Images目录"""
    try:
        from pathlib import Path

        # 检查是否存在“盖章页文件.pdf”
        images_dir = file_manager.get_func2_dir("images")
        pdf_file = images_dir / "盖章页文件.pdf"

        if not pdf_file.exists():
            return jsonify({"success": False, "error": f"PDF文件不存在: {pdf_file}"})

        # 从PDF提取图片
        extracted_images = stamp_service._extract_image_from_stamp(pdf_file, images_dir)

        return jsonify(
            {
                "success": True,
                "message": "图片提取完成",
                "count": len(extracted_images),
                "files": [str(file) for file in extracted_images],
            }
        )
    except Exception as e:
        logger.error(f"从PDF提取图片失败: {str(e)}", exc_info=True)
        return jsonify({"success": False, "error": f"从PDF提取图片失败: {str(e)}"})


@app.route("/api/start-stamp-overlay", methods=["POST"])
def start_stamp_overlay():
    """开始盖章页覆盖处理"""
    try:
        # 获取当前配置
        from GaiZhangYe.core.data_communication import get_data_service
        from pathlib import Path

        data = get_data_service().get_func2_data()

        # 执行盖章覆盖
        # 获取图片文件列表
        images_dir = file_manager.get_func2_dir("images")
        import os

        image_files = []
        if images_dir.exists():
            for filename in os.listdir(images_dir):
                if filename.endswith((".png", ".jpg", ".jpeg", ".bmp")):
                    image_files.append(Path(images_dir) / filename)

        result_files = stamp_service.run(
            configs=data.get("config", {}), image_files=image_files
        )

        return jsonify(
            {
                "success": True,
                "message": "盖章页覆盖完成",
                "files": [str(file) for file in result_files],
            }
        )
    except Exception as e:
        logger.error(f"盖章页覆盖失败: {str(e)}", exc_info=True)
        return jsonify({"success": False, "error": f"盖章页覆盖失败: {str(e)}"})


import sys
import threading
import os


@app.route("/api/shutdown", methods=["POST"])
def shutdown():
    """终止服务的API端点"""
    # 只能在开发环境下使用
    try:
        # 使用单独的线程来终止服务，确保能够返回响应
        def terminate_service():
            # 给响应时间，然后终止
            import time

            time.sleep(0.5)

            # 不同平台的终止方式
            try:
                if sys.platform == "win32":
                    # Windows系统
                    os.system(f"taskkill /F /PID {os.getpid()}")
                else:
                    # Linux/macOS系统
                    import signal

                    os.kill(os.getpid(), signal.SIGINT)
            except Exception:
                # 如果上述方法失败，尝试直接退出
                sys.exit()

        # 启动终止线程
        threading.Thread(target=terminate_service, daemon=True).start()

        # 返回成功响应
        return jsonify({"success": True, "message": "服务正在终止..."})
    except Exception as e:
        logger.error(f"终止服务失败: {e}")
        return jsonify({"success": False, "error": f"终止服务失败: {str(e)}"})


if __name__ == "__main__":
    logger.info("启动HTML可视化服务...")
    logger.info("服务将在 http://localhost:5001 启动")
    logger.info("按 Ctrl+C 停止服务")

    app.run(host="0.0.0.0", port=5001, debug=False)
