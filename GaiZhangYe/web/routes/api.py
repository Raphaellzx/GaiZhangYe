import os
import sys
from pathlib import Path
from flask import request, jsonify, current_app
from flask_restx import Namespace, Resource, fields

api_ns = Namespace('', description='API接口')

from GaiZhangYe.core.basic.file_manager import get_file_manager
from GaiZhangYe.core.basic.file_manager import (
    sort_files_windows_style,
    sort_dicts_by_name_windows_style,
)

# Initialize singletons used by APIs
file_manager = get_file_manager()


# 响应模型
response_model = api_ns.model('Response', {
    'success': fields.Boolean(required=True, description='请求是否成功'),
    'message': fields.String(description='响应消息'),
    'error': fields.String(description='错误信息'),
    'data': fields.Raw(description='响应数据'),
})

# 文件模型
file_model = api_ns.model('File', {
    'name': fields.String(required=True, description='文件名'),
    'stem': fields.String(description='文件名(不含扩展名)'),
    'page_count': fields.Integer(description='页数'),
    'total_pages': fields.Integer(description='总页数'),
})

# 目录模型
directory_model = api_ns.model('Directory', {
    'nostamped_word': fields.String(description='未盖章Word文件目录'),
    'nostamped_pdf': fields.String(description='未盖章PDF文件目录'),
    'stamped_pages': fields.String(description='已盖章页面目录'),
    'images': fields.String(description='图片目录'),
    'target_files': fields.String(description='目标文件目录'),
    'result_word': fields.String(description='结果Word文件目录'),
    'result_pdf': fields.String(description='结果PDF文件目录'),
})


@api_ns.route('/session-id')
class SessionId(Resource):
    @api_ns.doc('获取会话ID')
    @api_ns.response(200, '成功获取会话ID')
    def get(self):
        """获取会话ID"""
        return jsonify({"success": True, "session_id": current_app.config.get('APP_SESSION_ID')})


@api_ns.route('/scan-folder', methods=['POST'])
class ScanFolder(Resource):
    @api_ns.doc('扫描文件夹中的Word文件')
    @api_ns.expect(api_ns.model('ScanFolderRequest', {
        'path': fields.String(required=True, description='文件夹路径')
    }))
    @api_ns.response(200, '成功扫描到Word文件')
    @api_ns.response(400, '无效的文件夹路径')
    @api_ns.marshal_with(response_model)
    def post(self):
        """扫描文件夹中的Word文件"""
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


@api_ns.route('/scan-folder-with-images', methods=['POST'])
class ScanFolderWithImages(Resource):
    @api_ns.doc('扫描文件夹中的Word文件和图片')
    @api_ns.expect(api_ns.model('ScanFolderWithImagesRequest', {
        'word_path': fields.String(description='Word文件路径'),
        'image_path': fields.String(description='图片文件路径')
    }))
    @api_ns.response(200, '成功扫描到文件')
    @api_ns.marshal_with(response_model)
    def post(self):
        """扫描文件夹中的Word文件和图片"""
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


@api_ns.route('/word-to-pdf', methods=['POST'])
class WordToPdf(Resource):
    @api_ns.doc('Word转PDF')
    @api_ns.expect(api_ns.model('WordToPdfRequest', {
        'input_dir': fields.String(description='输入目录'),
        'output_dir': fields.String(description='输出目录')
    }))
    @api_ns.response(200, '成功转换为PDF')
    @api_ns.marshal_with(response_model)
    def post(self):
        """Word转PDF"""
        try:
            from GaiZhangYe.core.batch_convert import BatchConvertService

            data = request.get_json() or {}
            # 处理自定义路径
            if data.get('input_dir'):
                file_manager.set_custom_func2_dir('target_files', Path(data.get('input_dir')))
            if data.get('output_dir'):
                file_manager.set_custom_func2_dir('result_pdf', Path(data.get('output_dir')))

            input_dir = file_manager.get_func2_dir('target_files')
            output_dir = file_manager.get_func2_dir('result_pdf')

            convert_service = BatchConvertService()
            result_files = convert_service.run(input_dir, output_dir)
            result_files_str = [str(f) for f in result_files]
            return jsonify({"success": True, "message": f"转换完成！共生成 {len(result_files_str)} 个PDF文件", "output_dir": str(output_dir), "files": result_files_str})
        except Exception as e:
            current_app.logger.error(f"Word转PDF失败: {str(e)}", exc_info=True)
            return jsonify({"success": False, "error": f"转换失败: {str(e)}"})


@api_ns.route('/status')
class Status(Resource):
    @api_ns.doc('获取服务状态')
    @api_ns.response(200, '成功获取服务状态')
    def get(self):
        """获取服务状态"""
        return jsonify({"status": "running", "message": "盖章页工具HTML服务已启动"})


@api_ns.route('/open-directory')
class OpenDirectory(Resource):
    @api_ns.doc('打开目录')
    @api_ns.expect(api_ns.model('OpenDirectoryRequest', {
        'dir_name': fields.String(required=True, description='目录名称')
    }))
    @api_ns.response(200, '成功打开目录')
    @api_ns.response(400, '无效的目录名')
    @api_ns.marshal_with(response_model)
    def get(self):
        """打开目录"""
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
                os.startfile(dir_str)
                return jsonify({"success": True, "message": "目录已打开"})
            return jsonify({"success": False, "error": f"目录不存在: {dir_name}"})
        except Exception as e:
            current_app.logger.error(f"打开目录失败: {str(e)}")
            return jsonify({"success": False, "error": f"打开目录失败: {str(e)}"})


@api_ns.route('/get-default-output-path')
class GetDefaultOutputPath(Resource):
    @api_ns.doc('获取默认输出路径')
    @api_ns.response(200, '成功获取默认输出路径')
    @api_ns.marshal_with(response_model)
    def get(self):
        """获取默认输出路径"""
        try:
            if sys.platform == 'win32':
                documents_path = Path.home() / 'Documents'
            else:
                documents_path = Path.home() / 'Documents'
            return jsonify({"success": True, "path": str(documents_path)})
        except Exception as e:
            current_app.logger.error(f"获取默认输出路径失败: {str(e)}")
            return jsonify({"success": False, "error": str(e)})


@api_ns.route('/get-default-output-paths')
class GetDefaultOutputPaths(Resource):
    @api_ns.doc('获取默认输出路径')
    @api_ns.response(200, '成功获取默认输出路径')
    @api_ns.marshal_with(response_model)
    def get(self):
        """获取默认输出路径"""
        try:
            result_word_path = file_manager.get_func2_dir('result_word')
            result_pdf_path = file_manager.get_func2_dir('result_pdf')
            return jsonify({"success": True, "result_word_path": str(result_word_path), "result_pdf_path": str(result_pdf_path)})
        except Exception as e:
            current_app.logger.error(f"获取默认输出路径失败: {str(e)}")
            return jsonify({"success": False, "error": str(e)})


@api_ns.route('/prepare-stamp', methods=['POST'])
class PrepareStamp(Resource):
    @api_ns.doc('准备盖章页')
    @api_ns.expect(api_ns.model('PrepareStampRequest', {
        'target_pages': fields.String(required=True, description='目标页面'),
        'output_path': fields.String(required=True, description='输出路径'),
        'word_dir': fields.String(description='Word文件目录')
    }))
    @api_ns.response(200, '成功准备盖章页')
    @api_ns.response(400, '缺少必填参数')
    @api_ns.marshal_with(response_model)
    def post(self):
        """准备盖章页"""
        try:
            from GaiZhangYe.core.stamp_prepare import StampPrepareService

            data = request.get_json() or {}
            target_pages = data.get('target_pages')
            output_path = data.get('output_path')
            word_dir = data.get('word_dir')

            if not target_pages:
                return jsonify({"success": False, "error": "没有提供页面范围"})
            if not output_path:
                return jsonify({"success": False, "error": "没有提供输出路径"})

            # 使用file_manager管理自定义路径
            if word_dir:
                file_manager.set_custom_func1_dir("nostamped_word", Path(word_dir))
            if output_path:
                file_manager.set_custom_func1_dir("stamped_pages", Path(output_path))

            stamp_service = StampPrepareService()
            result_files = stamp_service.run(target_pages)
            return jsonify({"success": True, "message": "盖章页准备完成", "files": [str(f) for f in result_files]})
        except Exception as e:
            current_app.logger.error(f"准备盖章页失败: {str(e)}", exc_info=True)
            return jsonify({"success": False, "error": f"准备盖章页失败: {str(e)}"})


@api_ns.route('/directories')
class GetDirectories(Resource):
    @api_ns.doc('获取目录信息')
    @api_ns.response(200, '成功获取目录信息')
    def get(self):
        """获取目录信息"""
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


@api_ns.route('/refresh-data', methods=['POST'])
class RefreshData(Resource):
    @api_ns.doc('刷新数据')
    @api_ns.response(200, '成功刷新数据')
    @api_ns.marshal_with(response_model)
    def post(self):
        """刷新数据"""
        try:
            from GaiZhangYe.core.data_communication import get_data_service
            data_service = get_data_service()
            if data_service.scan_business_data():
                return jsonify({"success": True, "message": "数据文件已重新生成"})
            return jsonify({"success": False, "error": "数据文件重新生成失败"})
        except Exception as e:
            return jsonify({"success": False, "error": str(e)})


@api_ns.route('/func1/data', methods=['GET', 'POST'])
class Func1Data(Resource):
    @api_ns.doc('获取功能1数据')
    @api_ns.response(200, '成功获取功能1数据')
    @api_ns.marshal_with(response_model)
    def get(self):
        """获取功能1数据"""
        try:
            from GaiZhangYe.core.data_communication import get_data_service
            data = get_data_service().get_func1_data()
            # 转换为可序列化的字典格式
            data_dict = data.__dict__ if hasattr(data, '__dict__') else data
            return jsonify({"success": True, "data": data_dict})
        except Exception as e:
            current_app.logger.error(f"获取功能1数据失败: {str(e)}", exc_info=True)
            return jsonify({"success": False, "error": str(e)})

    @api_ns.doc('保存功能1数据')
    @api_ns.expect(api_ns.model('Func1DataRequest', {
        'target_pages': fields.Raw(required=True, description='目标页面配置')
    }))
    @api_ns.response(200, '成功保存功能1数据')
    @api_ns.marshal_with(response_model)
    def post(self):
        """保存功能1数据"""
        try:
            from GaiZhangYe.core.data_communication import get_data_service
            from GaiZhangYe.core.models.data_models import Func1Data

            data = request.get_json() or {}

            # 验证数据格式
            if 'target_pages' not in data:
                return jsonify({"success": False, "error": "缺少target_pages字段"})

            # 转换为Func1Data对象
            func1_data = Func1Data(target_pages=data['target_pages'])

            if get_data_service().save_func1_data(func1_data):
                return jsonify({"success": True, "message": "功能1数据保存成功"})
            return jsonify({"success": False, "error": "保存失败"})
        except Exception as e:
            current_app.logger.error(f"保存功能1数据失败: {str(e)}", exc_info=True)
            return jsonify({"success": False, "error": str(e)})


@api_ns.route('/func2/data', methods=['GET', 'POST'])
class Func2Data(Resource):
    @api_ns.doc('获取功能2数据')
    @api_ns.response(200, '成功获取功能2数据')
    @api_ns.marshal_with(response_model)
    def get(self):
        """获取功能2数据"""
        try:
            from GaiZhangYe.core.data_communication import get_data_service
            data = get_data_service().get_func2_data()
            # 转换为可序列化的字典格式
            data_dict = data.__dict__ if hasattr(data, '__dict__') else data
            return jsonify({"success": True, "data": data_dict})
        except Exception as e:
            current_app.logger.error(f"获取功能2数据失败: {str(e)}", exc_info=True)
            return jsonify({"success": False, "error": str(e)})

    @api_ns.doc('保存功能2数据')
    @api_ns.expect(api_ns.model('Func2DataRequest', {
        'configs': fields.List(fields.Raw, required=True, description='盖章配置列表')
    }))
    @api_ns.response(200, '成功保存功能2数据')
    @api_ns.marshal_with(response_model)
    def post(self):
        """保存功能2数据"""
        try:
            from GaiZhangYe.core.data_communication import get_data_service
            from GaiZhangYe.core.models.data_models import Func2Data, StampConfig

            data = request.get_json() or {}

            # 验证数据格式
            if 'configs' not in data:
                return jsonify({"success": False, "error": "缺少configs字段"})

            # 转换为Func2Data对象
            configs = []
            for config_data in data['configs']:
                configs.append(StampConfig(**config_data))

            func2_data = Func2Data(configs=configs)

            if get_data_service().save_func2_data(func2_data):
                return jsonify({"success": True, "message": "功能2数据保存成功"})
            return jsonify({"success": False, "error": "保存失败"})
        except Exception as e:
            current_app.logger.error(f"保存功能2数据失败: {str(e)}", exc_info=True)
            return jsonify({"success": False, "error": str(e)})


@api_ns.route('/extract-images-from-pdf', methods=['POST'])
class ExtractImagesFromPdf(Resource):
    @api_ns.doc('从PDF提取图片')
    @api_ns.response(200, '成功提取图片')
    @api_ns.response(400, 'PDF文件不存在')
    @api_ns.marshal_with(response_model)
    def post(self):
        """从PDF提取图片"""
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


@api_ns.route('/extract-pdf-images', methods=['POST'])
class ExtractPdfImages(Resource):
    @api_ns.doc('从指定PDF文件提取图片')
    @api_ns.expect(api_ns.model('ExtractPdfImagesRequest', {
        'pdf_path': fields.String(required=True, description='PDF文件路径')
    }))
    @api_ns.response(200, '成功提取图片')
    @api_ns.response(400, '无效的PDF文件')
    @api_ns.marshal_with(response_model)
    def post(self):
        """从指定PDF文件提取图片"""
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


@api_ns.route('/start-stamp-overlay', methods=['POST'])
class StartStampOverlay(Resource):
    @api_ns.doc('开始盖章页覆盖')
    @api_ns.expect(api_ns.model('StartStampOverlayRequest', {
        'target_word_dir': fields.String(required=True, description='目标Word文件目录'),
        'images_folder': fields.String(description='图片文件夹'),
        'result_word_path': fields.String(description='结果Word文件路径'),
        'result_pdf_path': fields.String(description='结果PDF文件路径')
    }))
    @api_ns.response(200, '成功开始盖章页覆盖')
    @api_ns.marshal_with(response_model)
    def post(self):
        """开始盖章页覆盖"""
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


@api_ns.route('/shutdown', methods=['POST'])
class Shutdown(Resource):
    @api_ns.doc('关闭服务')
    @api_ns.response(200, '成功关闭服务')
    @api_ns.marshal_with(response_model)
    def post(self):
        """关闭服务"""
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
