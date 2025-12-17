# GaiZhangYe/core/services/stamp_overlay.py
"""
功能2：盖章页覆盖服务
"""
from pathlib import Path
from typing import List
from GaiZhangYe.utils.logger import get_logger
from GaiZhangYe.core.basic.file_manager import FileManager
from GaiZhangYe.core.basic.file_processor import FileProcessor
from GaiZhangYe.core.basic.word_processor import WordProcessor
from GaiZhangYe.core.basic.pdf_processor import PdfProcessor
from GaiZhangYe.core.basic.image_processor import ImageProcessor
from GaiZhangYe.core.models.exceptions import BusinessError, ImageProcessError

logger = get_logger(__name__)


class StampOverlayService:
    """盖章页覆盖服务"""

    def __init__(self):
        self.file_manager = FileManager()
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
            image_width: int = None, image_files: List[Path] = None, configs=None) -> List[Path]:
        """
        执行功能2流程：
        1. 缩放图片后插入到目标 Word 文件
        2. 生成最终 Word/PDF
        :param stamp_file: 盖章后的PDF或图片文件
        :param target_word_dir: 目标Word文件目录
        :param image_width: 图片缩放宽度（默认配置中的默认值）
        :param image_files: 盖章图片文件列表，与Word文件一对一匹配
        :return: 生成的Word文件路径列表
        """
        logger.info("开始执行【功能2：盖章页覆盖】")
        try:
            # 1. 初始化目录
            images_dir, result_word_dir, result_pdf_dir, target_word_dir = self._init_directories(target_word_dir)

            # 2. 验证并准备图片文件
            self._validate_images(image_files)

            # 3. 获取并处理目标Word文件
            word_files = self._get_target_word_files(target_word_dir)

            # 4. 排序文件
            sorted_images = sorted(image_files, key=lambda x: x.name)
            sorted_word_files = sorted(word_files, key=lambda x: x.name)

            # 5. 批量插入图片并转换为PDF
            result_word_files = self._batch_insert_images_and_convert(
                sorted_images, sorted_word_files, images_dir, result_word_dir, result_pdf_dir, image_width, configs)

            logger.info(f"【功能2】执行完成，成功处理{len(result_word_files)}个Word文件")
            return result_word_files
        except Exception as e:
            logger.error("【功能2】执行失败", exc_info=True)
            raise BusinessError(f"盖章页覆盖失败：{str(e)}") from e

    def _init_directories(self, target_word_dir: Path) -> tuple:
        """初始化功能2所需的目录"""
        images_dir = self.file_manager.get_func2_dir("images")
        result_word_dir = self.file_manager.get_func2_dir("result_word")
        result_pdf_dir = self.file_manager.get_func2_dir("result_pdf")
        target_word_dir = target_word_dir or self.file_manager.get_func2_dir("target_files")
        return images_dir, result_word_dir, result_pdf_dir, target_word_dir

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

    def _batch_insert_images_and_convert(self, sorted_images: List[Path], sorted_word_files: List[Path],
                                         images_dir: Path, result_word_dir: Path, result_pdf_dir: Path,
                                         image_width: int, configs: list) -> List[Path]:
        """批量插入图片并将结果转换为PDF"""
        result_word_files = []

        for word in sorted_word_files:
            output_word = result_word_dir / f"{word.stem}.docx"

            # 检查是否有配置信息（UI层会传递此配置）
            current_config = self._get_current_config(word.name, configs)

            # 尝试使用配置模式处理
            if current_config:
                success = self._process_with_config(current_config, word, output_word, images_dir, image_width)
                if success:
                    result_word_files.append(output_word)

            # 如果配置模式处理失败或无配置，使用默认模式
            if not current_config or output_word not in result_word_files:
                success = self._process_with_default_mode(sorted_images, word, output_word, images_dir, image_width)
                if success:
                    result_word_files.append(output_word)

            # 将结果Word转换为PDF
            if output_word.exists():
                self._convert_word_to_pdf(output_word, result_pdf_dir)

        return result_word_files

    def _get_current_config(self, filename: str, configs: list) -> object:
        """根据文件名获取对应的配置信息"""
        if not configs:
            return None
        for config in configs:
            if config.filename == filename:
                return config
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

        # 插入每张图片
        for img_input, position in zip(current_config.image_files, current_config.insert_positions):
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

            # 获取插入位置
            image_location = position if isinstance(position, str) and position != "" else "last_page"

            # 插入图片
            self.word_processor.insert_image_to_word(temp_output, final_image, image_location, temp_output)

        # 将临时文件重命名为最终输出
        if os.path.exists(output_word):
            os.unlink(output_word)
        os.rename(temp_output, output_word)

        return True

    def _process_with_default_mode(self, sorted_images: List[Path], word: Path, output_word: Path,
                                  images_dir: Path, image_width: int) -> bool:
        """使用默认模式（文件名匹配）处理Word文件"""
        logger.info(f"[默认模式] 使用文件名匹配方式处理 Word 文件 {word.name}")

        # 查找对应的图片
        img_for_word = None
        for img in sorted_images:
            if img.stem == word.stem:
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

        # 插入图片
        self.word_processor.insert_image_to_word(word, final_image, "last_page", output_word)
        return True

    def _convert_word_to_pdf(self, output_word: Path, result_pdf_dir: Path) -> None:
        """将Word文件转换为PDF"""
        output_pdf = result_pdf_dir / f"{output_word.stem}_stamped.pdf"
        self.word_processor.word_to_pdf(output_word, output_pdf)
        logger.debug(f"生成最终PDF：{output_pdf}")