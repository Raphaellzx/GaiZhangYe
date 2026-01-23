# GaiZhangYe/core/services/stamp_overlay.py
"""
功能2：盖章页覆盖服务
"""
import re
from pathlib import Path
from typing import List
from GaiZhangYe.utils.logger import get_logger
from GaiZhangYe.core.basic.file_processor import windows_natural_sort_key
from GaiZhangYe.core.basic.file_manager import get_file_manager
from GaiZhangYe.core.basic.file_processor import FileProcessor
from GaiZhangYe.core.basic.word_processor import WordProcessor
from GaiZhangYe.core.basic.pdf_processor import PdfProcessor
from GaiZhangYe.core.basic.image_processor import ImageProcessor
from GaiZhangYe.core.models.exceptions import BusinessError
from GaiZhangYe.core.data_communication import get_data_service

logger = get_logger(__name__)


class StampOverlayService:
    """盖章页覆盖服务"""

    def __init__(self):
        self.file_manager = get_file_manager()
        self.file_processor = FileProcessor()
        self.word_processor = WordProcessor()
        self.pdf_processor = PdfProcessor()
        self.image_processor = ImageProcessor()

    def _extract_image_from_stamp(self, stamp_file: Path, output_dir: Path) -> List[Path]:
        """
        从盖章PDF中提取图片
        :param stamp_file: 盖章PDF文件路径
        :param output_dir: 图片输出目录
        :return: 提取的图片路径列表
        """
        return self.pdf_processor.extract_images(stamp_file, output_dir)

    def run(self, target_word_dir: Path = None,
            image_width: int = None, image_files: List[Path] = None, configs=None,
            result_word_dir: Path = None, result_pdf_dir: Path = None) -> List[Path]:
        """
        执行功能2流程：
        1. 缩放图片后插入到目标 Word 文件
        2. 生成最终 Word/PDF
        :param target_word_dir: 目标Word文件目录
        :param image_width: 图片缩放宽度（默认配置中的默认值）
        :param image_files: 盖章图片文件列表，与Word文件一对一匹配
        :param configs: 配置字典
        :param result_word_dir: 输出Word文件目录（可选）
        :param result_pdf_dir: 输出PDF文件目录（可选）
        :return: 生成的Word文件路径列表
        """
        logger.info("开始执行【功能2：盖章页覆盖】")
        try:
            # 1. 初始化目录
            images_dir, final_result_word_dir, final_result_pdf_dir, target_word_dir = self._init_directories(
                target_word_dir, result_word_dir, result_pdf_dir)

            # 2. 验证并准备图片文件
            # 如果没有提供configs，才需要验证image_files
            # configs 可能是 None 或者 空字典 {}
            has_valid_config = configs and isinstance(configs, dict) and len(configs) > 0
            
            if not has_valid_config:
                # 只在没有有效配置时验证image_files
                self._validate_images(image_files)
            else:
                logger.info("使用UI配置模式，无需提供image_files")

            # 3. 获取并处理目标Word文件
            word_files = self._get_target_word_files(target_word_dir)

            # 应用自然排序
            sorted_word_files = sorted(word_files, key=windows_natural_sort_key)

            # 根据情况生成sorted_images
            sorted_images = []
            # 无论是配置模式还是默认模式都需要sorted_images
            # 因为配置模式可能会失败回退到默认模式
            if image_files:
                sorted_images = sorted(image_files, key=windows_natural_sort_key)

            # 5. 批量插入图片并转换为PDF
            result_word_files = self._batch_insert_images_and_convert(
                sorted_word_files, images_dir, final_result_word_dir, final_result_pdf_dir, image_width, configs, sorted_images)

            logger.info(f"【功能2】执行完成，成功处理{len(result_word_files)}个Word文件")
            return result_word_files
        except Exception as e:
            logger.error("【功能2】执行失败", exc_info=True)
            raise BusinessError(f"盖章页覆盖失败：{str(e)}") from e

    def _init_directories(self, target_word_dir: Path = None, result_word_dir: Path = None, result_pdf_dir: Path = None) -> tuple:
        """初始化功能2所需的目录
        :param target_word_dir: 目标Word文件目录
        :param result_word_dir: 自定义输出Word目录
        :param result_pdf_dir: 自定义输出PDF目录
        """
        images_dir = self.file_manager.get_func2_dir("images")
        
        # 如果提供了自定义输出目录，使用自定义目录；否则使用默认目录
        final_result_word_dir = Path(result_word_dir) if result_word_dir else self.file_manager.get_func2_dir("result_word")
        final_result_pdf_dir = Path(result_pdf_dir) if result_pdf_dir else self.file_manager.get_func2_dir("result_pdf")
        
        # 创建输出目录，如果不存在的话
        final_result_word_dir.mkdir(parents=True, exist_ok=True)
        final_result_pdf_dir.mkdir(parents=True, exist_ok=True)
        
        target_word_dir = target_word_dir or self.file_manager.get_func2_dir("target_files")
        return images_dir, final_result_word_dir, final_result_pdf_dir, target_word_dir

    def _validate_images(self, image_files: List[Path]) -> None:
        """验证图片文件列表是否为空"""
        if not image_files or len(image_files) == 0:
            raise BusinessError("没有可处理的图片文件")

    def _get_target_word_files(self, target_word_dir: Path) -> List[Path]:
        """获取目标Word文件目录中的所有Word文件"""
        word_files = self.file_processor.list_files(target_word_dir, [".docx", ".doc"])
        if not word_files:
            raise BusinessError(f"目标Word目录{target_word_dir}中无Word文件")
        return word_files

    def _batch_insert_images_and_convert(self, sorted_word_files: List[Path],
                                         images_dir: Path, result_word_dir: Path, result_pdf_dir: Path,
                                         image_width: int, configs: dict, sorted_images: List[Path] = None) -> List[Path]:
        """批量插入图片并将结果转换为PDF
        核心逻辑：
        1. 将图片按顺序分配给Word文件
        2. 支持一个Word文件对应多张图片
        3. 使用UI层生成的配置控制图片插入位置
        示例：
        images_list = [img1, img2, img3, img4, img5]
        文件1需1张 → [img1]
        文件2需2张 → [img2, img3]
        文件3需2张 → [img4, img5]
        """
        result_word_files = []
        # 图片索引指针，用于按顺序分配图片
        image_index = 0
        total_images = len(sorted_images)

        for i, word in enumerate(sorted_word_files):
            output_word = result_word_dir / f"{word.stem}.docx"

            # 检查是否有配置信息（UI层会传递此配置）
            current_config = self._get_current_config(word.name, configs)
            processed_successfully = False

            if current_config:
                logger.info(f"[UI配置模式] 处理 Word 文件 {word.name}")

                try:
                    # 分配图片：将UI配置中实际存在的图片路径更新为实际的完整路径
                    actual_image_paths = []
                    for img_path_str in current_config.image_files:
                        img_path = Path(img_path_str)
                        # 支持两种格式：完整路径或仅文件名
                        if img_path.exists():
                            actual_image_path = img_path
                        else:
                            actual_image_path = images_dir / img_path_str
                            if not actual_image_path.exists():
                                logger.warning(f"配置中指定的图片不存在：{img_path_str}")
                                continue
                        actual_image_paths.append(str(actual_image_path))

                    # 更新当前配置的图片路径为实际完整路径
                    # 注：这里假设current_config是可变的或可以修改的
                    if hasattr(current_config, '__dict__'):
                        current_config.__dict__['image_files'] = actual_image_paths

                    # 使用配置模式处理
                    success = self._process_with_config(current_config, word, output_word, images_dir, image_width)
                    if success:
                        result_word_files.append(output_word)
                        processed_successfully = True

                        # 将结果Word转换为PDF
                        if output_word.exists():
                            self._convert_word_to_pdf(output_word, result_pdf_dir)
                            # 验证PDF是否生成（兼容带/不带 _stamped 后缀的命名）
                            pdf_file = self._find_pdf_file(result_pdf_dir, output_word.stem)
                            if not pdf_file or not pdf_file.exists():
                                logger.warning(f"PDF文件 {result_pdf_dir / output_word.stem} 未生成，将重新尝试一次")
                                self._convert_word_to_pdf(output_word, result_pdf_dir)

                        logger.info(f"[UI配置模式] 成功处理 Word 文件 {word.name}，插入图片 {len(actual_image_paths)} 张")
                except Exception as e:
                    logger.error(f"[UI配置模式] 处理失败 Word 文件 {word.name}：{str(e)}")

            # 如果配置模式处理失败或无配置，使用默认模式：1张图片/文件
            if not processed_successfully and image_index < total_images:
                logger.info(f"[默认模式] 处理 Word 文件 {word.name}")

                try:
                    # 使用配置模式处理单张图片的默认情况
                    import shutil
                    import os

                    # 计算当前Word的页数，作为默认插入页码
                    try:
                        default_page_for_word = self.word_processor.get_word_page_count(word)
                    except Exception:
                        default_page_for_word = 1

                    temp_config = type('TempConfig', (), {
                        'filename': word.name,
                        'image_files': [str(sorted_images[image_index])],
                        'insert_positions': [default_page_for_word]
                    })()
                    logger.info(f"[默认模式] 为 {word.name} 使用默认插入页码: {default_page_for_word}")

                    # 创建临时文件并处理
                    temp_output = output_word.parent / f"{word.stem}_temp.docx"
                    shutil.copy2(word, temp_output)

                    # 插入单张图片
                    success = self._process_with_config(temp_config, word, output_word, images_dir, image_width)

                    if success:
                        result_word_files.append(output_word)
                        image_index += 1

                        # 将结果Word转换为PDF
                            if output_word.exists():
                            self._convert_word_to_pdf(output_word, result_pdf_dir)
                            # 验证PDF是否生成（兼容带/不带 _stamped 后缀的命名）
                            pdf_file = self._find_pdf_file(result_pdf_dir, output_word.stem)
                            if not pdf_file or not pdf_file.exists():
                                logger.warning(f"PDF文件 {result_pdf_dir / output_word.stem} 未生成，将重新尝试一次")
                                self._convert_word_to_pdf(output_word, result_pdf_dir)

                        logger.info(f"[默认模式] 成功处理 Word 文件 {word.name}，插入图片 1 张")
                        processed_successfully = True
                except Exception as e:
                    logger.error(f"[默认模式] 处理失败 Word 文件 {word.name}：{str(e)}")

            if not processed_successfully:
                if current_config:
                    logger.warning(f"已无足够图片或配置错误，无法处理 Word 文件 {word.name}")
                else:
                    logger.warning(f"已无足够图片，无法处理 Word 文件 {word.name}")

        # 安全检查：确保所有成功处理的Word文件都已转换为PDF
        for word_file in result_word_files:
            if word_file.exists():
                # 检查PDF是否已存在（兼容带/不带 _stamped 后缀的命名）
                pdf_file = self._find_pdf_file(result_pdf_dir, word_file.stem)

                if not pdf_file or not pdf_file.exists():
                    logger.info(f"【安全检查】重新生成 PDF 文件：{result_pdf_dir / word_file.stem}")
                    self._convert_word_to_pdf(word_file, result_pdf_dir)

        return result_word_files

    def _get_current_config(self, filename: str, configs: dict) -> object:
        """根据文件名获取对应的配置信息"""
        if not configs:
            return None
        if filename in configs:
            # 将字典结构转换为对象结构以便后续处理
            import collections
            Config = collections.namedtuple('Config', ['filename', 'image_files', 'insert_positions'])
            items = configs[filename]

            # 处理用户输入-1的情况
            # 当items是字符串"-1"或不包含image/position键的结构时，返回None使用默认模式
            try:
                # 验证items是否是一个包含image和position的字典列表
                valid_items = []
                for item in items:
                    if isinstance(item, dict) and 'image' in item and 'position' in item:
                        valid_items.append(item)
                    else:
                        # 包含无效项，直接使用默认模式
                        return None

                # 如果没有有效配置项，返回None
                if not valid_items:
                    return None

                return Config(
                    filename=filename,
                    image_files=[item['image'] for item in valid_items],
                    insert_positions=[item['position'] for item in valid_items]
                )
            except Exception:
                # 当items不是预期的列表结构时，返回None使用默认模式
                return None
        return None

    def _process_with_config(self, current_config: object, word: Path, output_word: Path,
                            images_dir: Path, image_width: int) -> bool:
        """使用配置模式处理Word文件"""
        logger.info(f"[UI配置模式] 处理 Word 文件 {word.name}")

        # 验证配置完整性
        if not hasattr(current_config, 'image_files') or not hasattr(current_config, 'insert_positions'):
            logger.warning(f"[UI配置模式] Word 文件 {word.name} 的配置不完整，将使用默认处理方式")
            return False
        elif not current_config.image_files or not current_config.insert_positions:
            logger.warning(f"[UI配置模式] Word 文件 {word.name} 的配置缺少图片或位置信息，将使用默认处理方式")
            return False

        # 确保图片文件和插入位置数量一致
        if len(current_config.image_files) != len(current_config.insert_positions):
            logger.warning(f"[UI配置模式] Word 文件 {word.name} 的图片数量和位置数量不一致，将使用默认处理方式")
            return False

        logger.info(f"[UI配置模式] Word 文件 {word.name} 将插入 {len(current_config.image_files)} 张图片")

        import shutil
        import os

        # 插入所有图片到同一个Word文件
        temp_output = output_word.parent / f"{word.stem}_temp.docx"
        shutil.copy2(word, temp_output)

        # 规范化插入位置：将 'last_page' 或 非数值项回退为文档总页数（后端强制处理旧配置）
        def _normalize_positions(positions):
            try:
                total_pages = self.word_processor.get_word_page_count(word)
            except Exception:
                total_pages = 1

            normalized = []
            for pos in positions:
                try:
                    if isinstance(pos, str):
                        s = pos.strip().lower()
                        if s == 'last_page' or s == '-1':
                            n = total_pages
                        else:
                            n = int(float(s))
                    else:
                        n = int(pos)
                except Exception:
                    n = total_pages

                # Clamp to [1, total_pages]
                if n < 1:
                    n = 1
                if total_pages and n > total_pages:
                    n = total_pages
                normalized.append(n)

            return normalized

        normalized_positions = _normalize_positions(current_config.insert_positions)

        # 插入每张图片
        for img_input, position in zip(current_config.image_files, normalized_positions):
            logger.info(f"[UI配置模式] 将图片 {img_input} 插入文件 {word.name} 的页码 {position}")
            # 支持两种图片路径格式：直接路径和文件名
            img_path = Path(img_input) if Path(img_input).exists() else images_dir / img_input

            # 验证图片文件存在
            if not img_path.exists():
                logger.warning(f"图片文件不存在：{img_path}，跳过该图片")
                continue

            # 缩放图片（如果需要）
            final_image = img_path
            if image_width:
                scaled_image = images_dir / f"scaled_{img_path.name}"
                self.image_processor.resize_image(img_path, scaled_image,
                                                target_width=image_width, keep_ratio=True)
                final_image = scaled_image

            # 获取插入位置（仅数值）
            try:
                # position 可能已经是数字，也可能是字符串数字
                image_page = int(position)
            except Exception:
                try:
                    image_page = self.word_processor.get_word_page_count(word)
                except Exception:
                    image_page = 1

            # 插入图片（传递数值页码）
            self.word_processor.insert_image_to_word(temp_output, final_image, image_page, temp_output)

        # 将临时文件重命名为最终输出
        if os.path.exists(output_word):
            os.unlink(output_word)
        os.rename(temp_output, output_word)

        return True

    def _process_with_default_mode(self, sorted_images: List[Path], word: Path, output_word: Path,
                                  images_dir: Path, image_width: int, used_images: set) -> bool:
        """使用默认模式（文件名匹配）处理Word文件"""
        logger.info(f"[默认模式] 使用文件名匹配方式处理 Word 文件 {word.name}")

        # 查找对应的图片
        img_for_word = None
        for img in sorted_images:
            if img.stem == word.stem and str(img) not in used_images:
                img_for_word = img
                break

        # 如果没有找到严格匹配的图片，尝试按顺序分配
        if not img_for_word:
            logger.info(f"[默认模式] 未找到严格匹配的图片，尝试按顺序分配")
            for img in sorted_images:
                if str(img) not in used_images:
                    img_for_word = img
                    break

        if not img_for_word:
            logger.warning(f"未找到与Word文件 {word.name} 匹配的图片")
            return False

        # 缩放图片（如果需要）
        final_image = img_for_word
        if image_width:
            scaled_image = images_dir / f"scaled_{img_for_word.name}"
            self.image_processor.resize_image(img_for_word, scaled_image,
                                            target_width=image_width, keep_ratio=True)
            final_image = scaled_image

        # 插入图片，计算最后一页的页码并传入数值
        try:
            last_page = self.word_processor.get_word_page_count(word)
        except Exception:
            last_page = 1
        self.word_processor.insert_image_to_word(word, final_image, last_page, output_word)
        return True

    def _convert_word_to_pdf(self, output_word: Path, result_pdf_dir: Path) -> None:
        """将Word文件转换为PDF"""
        output_pdf = result_pdf_dir / f"{output_word.stem}"
        self.word_processor.word_to_pdf(output_word, output_pdf)
        logger.debug(f"生成最终PDF：{output_pdf}")

    def _find_pdf_file(self, result_pdf_dir: Path, stem: str) -> Path:
        """在结果目录中查找与给定word stem对应的PDF文件，兼容带或不带 "_stamped" 后缀的命名"""
        # 优先匹配常见命名
        candidates = []
        try:
            stamped = result_pdf_dir / f"{stem}_stamped.pdf"
            normal = result_pdf_dir / f"{stem}.pdf"
            if stamped.exists():
                return stamped
            if normal.exists():
                return normal

            # 退而求其次：查找没有扩展名或其他变体（e.g., 导出时未带 .pdf）
            for p in result_pdf_dir.iterdir():
                if not p.is_file():
                    continue
                if p.stem == stem or p.stem == f"{stem}_stamped":
                    candidates.append(p)
        except Exception:
            return None

        return candidates[0] if candidates else None