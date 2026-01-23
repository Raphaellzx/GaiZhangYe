import os
import sys
import json
import uuid
from pathlib import Path
from flask import Blueprint, request, jsonify, current_app

from GaiZhangYe.core.basic.file_manager import get_file_manager
from GaiZhangYe.core.basic.file_processor import (
    windows_natural_sort_key,
    sort_files_windows_style,
    sort_dicts_by_name_windows_style,
)

api_bp = Blueprint('api', __name__, url_prefix='/api')

# Initialize singletons used by APIs
file_manager = get_file_manager()


@api_bp.route('/session-id')
def session_id():
    return jsonify({"success": True, "session_id": current_app.config.get('APP_SESSION_ID')})


@api_bp.route('/scan-folder', methods=['POST'])
def scan_folder():
    try:
        from GaiZhangYe.core.basic.word_processor import WordProcessor

        data = request.get_json()
        folder_path = data.get('path')

        if not folder_path:
            return jsonify({"success": False, "error": "未提供文件夹路径"})

        folder_path = Path(folder_path)
        if not folder_path.exists() or not folder_path.is_dir():
            return jsonify({"success": False, "error": f"路径不存在或不是目录: {folder_path}"})

        word_processor = WordProcessor()
        word_files = []

        for filename in os.listdir(folder_path):
            if filename.endswith((".docx", ".doc")):
                file_path = folder_path / filename
                page_count = None
                try:
                    page_count = word_processor.get_word_page_count(file_path)
                except Exception as e:
                    current_app.logger.warning(f"获取文件 {filename} 页数失败: {e}")

                word_files.append({"name": filename, "stem": os.path.splitext(filename)[0], "page_count": page_count})

        word_files = sort_dicts_by_name_windows_style(word_files, 'name')
        return jsonify({"success": True, "files": word_files, "count": len(word_files)})
    except Exception as e:
        current_app.logger.error(f"扫描文件夹失败: {e}", exc_info=True)
        return jsonify({"success": False, "error": str(e)})


@api_bp.route('/scan-folder-with-images', methods=['POST'])
def scan_folder_with_images():
    try:
        from GaiZhangYe.core.basic.word_processor import WordProcessor

        data = request.get_json()
        word_folder = data.get('word_path')
        image_folder = data.get('image_path')

        word_files = []
        image_files = []

        if word_folder:
            word_folder = Path(word_folder)
            if word_folder.exists() and word_folder.is_dir():
                wp = WordProcessor()
                for filename in os.listdir(word_folder):
                    if filename.endswith((".docx", ".doc")):
                        file_path = word_folder / filename
                        try:
                            total_pages = wp.get_word_page_count(file_path)
                            total_pages = total_pages if total_pages > 0 else 1
                        except Exception as e:
                            current_app.logger.warning(f"获取文件页数失败 {filename}: {e}")
                            total_pages = 1
                        word_files.append({"name": filename, "total_pages": total_pages})
                word_files = sort_dicts_by_name_windows_style(word_files, 'name')

        if image_folder:
            image_folder = Path(image_folder)
            if image_folder.exists() and image_folder.is_dir():
                for filename in os.listdir(image_folder):
                    if filename.endswith((".png", ".jpg", ".jpeg", ".bmp")):
                        image_files.append(filename)
                image_files = sort_files_windows_style(image_files)

        return jsonify({
            "success": True,
            "word_files": word_files,
            "image_files": image_files,
            "word_count": len(word_files),
            "image_count": len(image_files),
        })
    except Exception as e:
        current_app.logger.error(f"扫描文件夹失败: {e}", exc_info=True)
        return jsonify({"success": False, "error": str(e)})


@api_bp.route('/word-to-pdf', methods=['POST'])
def api_word_to_pdf():
    try:
        from GaiZhangYe.core.batch_convert import BatchConvertService

        data = request.get_json() or {}
        input_dir = Path(data.get('input_dir')) if data.get('input_dir') else file_manager.get_func2_dir('target_files')
        output_dir = Path(data.get('output_dir')) if data.get('output_dir') else file_manager.get_func2_dir('result_pdf')

        convert_service = BatchConvertService()
        result_files = convert_service.run(input_dir, output_dir)
        result_files_str = [str(f) for f in result_files]
        return jsonify({"success": True, "message": f"转换完成！共生成 {len(result_files_str)} 个PDF文件", "output_dir": str(output_dir), "files": result_files_str})
    except Exception as e:
        current_app.logger.error(f"Word转PDF失败: {str(e)}", exc_info=True)
        return jsonify({"success": False, "error": f"转换失败: {str(e)}"})


@api_bp.route('/status')
def status():
    return jsonify({"status": "running", "message": "盖章页工具HTML服务已启动"})


@api_bp.route('/open-directory')
def open_directory():
    dir_name = request.args.get('dir_name')
    try:
        if not dir_name:
            return jsonify({"success": False, "error": "未提供目录名"})

        func1_map = {"Nostamped_Word": "nostamped_word", "Nostamped_PDF": "nostamped_pdf", "Stamped_Pages": "stamped_pages"}
        func2_map = {"Images": "images", "TargetFiles": "target_files", "Result_Word": "result_word", "Result_PDF": "result_pdf"}

        if dir_name in func1_map:
            dir_path = file_manager.get_func1_dir(func1_map[dir_name])
        elif dir_name in func2_map:
            dir_path = file_manager.get_func2_dir(func2_map[dir_name])
        else:
            return jsonify({"success": False, "error": f"无效的目录名: {dir_name}"})

        if dir_path.exists():
            dir_str = str(dir_path)
            if sys.platform == 'win32':
                os.startfile(dir_str)
            elif sys.platform == 'darwin':
                os.system(f"open '{dir_str}'")
            else:
                os.system(f"xdg-open '{dir_str}'")
            return jsonify({"success": True, "message": "目录已打开"})
        return jsonify({"success": False, "error": f"目录不存在: {dir_name}"})
    except Exception as e:
        current_app.logger.error(f"打开目录失败: {str(e)}")
        return jsonify({"success": False, "error": f"打开目录失败: {str(e)}"})


@api_bp.route('/get-default-output-path')
def get_default_output_path():
    try:
        if sys.platform == 'win32':
            documents_path = Path.home() / 'Documents'
        else:
            documents_path = Path.home() / 'Documents'
        return jsonify({"success": True, "path": str(documents_path)})
    except Exception as e:
        current_app.logger.error(f"获取默认输出路径失败: {str(e)}")
        return jsonify({"success": False, "error": str(e)})


@api_bp.route('/get-default-output-paths')
def get_default_output_paths():
    try:
        result_word_path = file_manager.get_func2_dir('result_word')
        result_pdf_path = file_manager.get_func2_dir('result_pdf')
        return jsonify({"success": True, "result_word_path": str(result_word_path), "result_pdf_path": str(result_pdf_path)})
    except Exception as e:
        current_app.logger.error(f"获取默认输出路径失败: {str(e)}")
        return jsonify({"success": False, "error": str(e)})


@api_bp.route('/prepare-stamp', methods=['POST'])
def prepare_stamp():
    try:
        from GaiZhangYe.core.stamp_prepare import StampPrepareService

        data = request.get_json() or {}
        target_pages = data.get('target_pages')
        output_path = data.get('output_path')
        # optional custom input Word directory for Nostamped_Word
        word_dir = data.get('word_dir')

        if not target_pages:
            return jsonify({"success": False, "error": "没有提供页面范围"})
        if not output_path:
            return jsonify({"success": False, "error": "没有提供输出路径"})

        stamp_service = StampPrepareService()
        # If a custom word_dir was provided, pass it through; otherwise use default configured directory
        result_files = stamp_service.run(target_pages, word_dir=Path(word_dir) if word_dir else None, output_dir=Path(output_path))
        return jsonify({"success": True, "message": "盖章页准备完成", "files": [str(f) for f in result_files]})
    except Exception as e:
        current_app.logger.error(f"准备盖章页失败: {str(e)}", exc_info=True)
        return jsonify({"success": False, "error": f"准备盖章页失败: {str(e)}"})


@api_bp.route('/directories')
def get_directories():
    try:
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
        current_app.logger.error(f"获取目录信息失败: {e}")
        return jsonify({"error": str(e)}), 500


@api_bp.route('/refresh-data', methods=['POST'])
def api_refresh_data():
    try:
        from GaiZhangYe.core.data_communication import get_data_service
        data_service = get_data_service()
        if data_service.scan_business_data():
            return jsonify({"success": True, "message": "数据文件已重新生成"})
        return jsonify({"success": False, "error": "数据文件重新生成失败"})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@api_bp.route('/func1/data', methods=['GET'])
def api_get_func1_data():
    try:
        from GaiZhangYe.core.data_communication import get_data_service
        data = get_data_service().get_func1_data()
        return jsonify({"success": True, "data": data})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@api_bp.route('/func1/data', methods=['POST'])
def api_save_func1_data():
    try:
        from GaiZhangYe.core.data_communication import get_data_service
        data = request.get_json() or {}
        if get_data_service().save_func1_data(data):
            return jsonify({"success": True})
        return jsonify({"success": False, "error": "保存失败"})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@api_bp.route('/func2/data', methods=['GET'])
def api_get_func2_data():
    try:
        from GaiZhangYe.core.data_communication import get_data_service
        data = get_data_service().get_func2_data()
        return jsonify({"success": True, "data": data})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@api_bp.route('/func2/data', methods=['POST'])
def api_save_func2_data():
    try:
        from GaiZhangYe.core.data_communication import get_data_service
        data = request.get_json() or {}
        if get_data_service().save_func2_data(data):
            return jsonify({"success": True})
        return jsonify({"success": False, "error": "保存失败"})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@api_bp.route('/extract-images-from-pdf', methods=['POST'])
def extract_images_from_pdf():
    try:
        images_dir = file_manager.get_func2_dir('images')
        pdf_file = images_dir / '盖章页文件.pdf'
        if not pdf_file.exists():
            return jsonify({"success": False, "error": f"PDF文件不存在: {pdf_file}"})

        from GaiZhangYe.core.stamp_overlay import StampOverlayService
        stamp_service = StampOverlayService()
        extracted_images = stamp_service._extract_image_from_stamp(pdf_file, images_dir)

        return jsonify({"success": True, "message": "图片提取完成", "count": len(extracted_images), "files": [str(f) for f in extracted_images]})
    except Exception as e:
        current_app.logger.error(f"从PDF提取图片失败: {str(e)}", exc_info=True)
        return jsonify({"success": False, "error": f"从PDF提取图片失败: {str(e)}"})


@api_bp.route('/extract-pdf-images', methods=['POST'])
def extract_pdf_images():
    try:
        data = request.get_json() or {}
        pdf_path_str = data.get('pdf_path', '').strip()
        if not pdf_path_str:
            return jsonify({"success": False, "error": "请提供PDF文件路径"})
        pdf_path = Path(pdf_path_str)
        if not pdf_path.exists() or not pdf_path.is_file() or pdf_path.suffix.lower() != '.pdf':
            return jsonify({"success": False, "error": f"无效的PDF文件: {pdf_path}"})

        output_folder = pdf_path.parent / f"{pdf_path.stem}_images"
        output_folder.mkdir(parents=True, exist_ok=True)

        from GaiZhangYe.core.basic.pdf_processor import PdfProcessor
        pdf_processor = PdfProcessor()
        extracted_images = pdf_processor.extract_images(pdf_path, output_folder)
        image_files = [img.name for img in extracted_images]

        return jsonify({"success": True, "message": f"成功从PDF提取 {len(image_files)} 张图片", "output_folder": str(output_folder), "image_files": image_files, "count": len(image_files)})
    except Exception as e:
        current_app.logger.error(f"从PDF提取图片失败: {str(e)}", exc_info=True)
        return jsonify({"success": False, "error": f"从PDF提取图片失败: {str(e)}"})


@api_bp.route('/start-stamp-overlay', methods=['POST'])
def start_stamp_overlay():
    try:
        from GaiZhangYe.core.data_communication import get_data_service

        data = request.get_json() or {}
        target_word_dir = data.get('target_word_dir')
        images_folder = data.get('images_folder')
        result_word_path = data.get('result_word_path')
        result_pdf_path = data.get('result_pdf_path')

        if not target_word_dir:
            return jsonify({"success": False, "error": "未提供Word文件夹路径"})

        target_word_dir = Path(target_word_dir)
        if not target_word_dir.exists() or not target_word_dir.is_dir():
            return jsonify({"success": False, "error": f"Word文件夹不存在: {target_word_dir}"})

        config_data = get_data_service().get_func2_data()
        has_config = config_data and config_data.get('config') and len(config_data.get('config', {})) > 0

        image_files = []
        if images_folder:
            images_dir = Path(images_folder)
            if images_dir.exists() and images_dir.is_dir():
                for filename in os.listdir(images_dir):
                    if filename.endswith((".png", ".jpg", ".jpeg", ".bmp")):
                        image_files.append(images_dir / filename)
        else:
            images_dir = file_manager.get_func2_dir('images')
            if not has_config and images_dir.exists():
                for filename in os.listdir(images_dir):
                    if filename.endswith((".png", ".jpg", ".jpeg", ".bmp")):
                        image_files.append(Path(images_dir) / filename)

        from GaiZhangYe.core.stamp_overlay import StampOverlayService
        stamp_service = StampOverlayService()
        result_files = stamp_service.run(
            target_word_dir=target_word_dir,
            configs=config_data.get('config', {}),
            image_files=image_files if image_files else None,
            result_word_dir=Path(result_word_path) if result_word_path else None,
            result_pdf_dir=Path(result_pdf_path) if result_pdf_path else None,
        )

        return jsonify({"success": True, "message": "盖章页覆盖完成", "files": [str(f) for f in result_files]})
    except Exception as e:
        current_app.logger.error(f"盖章页覆盖失败: {str(e)}", exc_info=True)
        return jsonify({"success": False, "error": f"盖章页覆盖失败: {str(e)}"})


@api_bp.route('/shutdown', methods=['POST'])
def shutdown():
    try:
        import threading

        def terminate_service():
            import time
            time.sleep(0.5)
            try:
                if sys.platform == 'win32':
                    os.system(f"taskkill /F /PID {os.getpid()}")
                else:
                    import signal
                    os.kill(os.getpid(), signal.SIGINT)
            except Exception:
                sys.exit()

        threading.Thread(target=terminate_service, daemon=True).start()
        return jsonify({"success": True, "message": "服务正在终止..."})
    except Exception as e:
        current_app.logger.error(f"终止服务失败: {e}")
        return jsonify({"success": False, "error": f"终止服务失败: {str(e)}"})
