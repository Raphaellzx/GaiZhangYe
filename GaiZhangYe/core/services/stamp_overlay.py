# GaiZhangYe/core/services/stamp_overlay.py
"""
功能2：盖章页覆盖服务
"""
from pathlib import Path
from typing import List
from GaiZhangYe.utils.logger import get_logger
from GaiZhangYe.core.file_manager import FileManager
from GaiZhangYe.core.file_processor import FileProcessor
from GaiZhangYe.core.word_processor import WordProcessor
from GaiZhangYe.core.pdf_processor import PdfProcessor
from GaiZhangYe.core.image_processor import ImageProcessor
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

    def run(self, stamp_file: Path, target_word_dir: Path = None,
            image_width: int = None) -> List[Path]:
        """
        执行功能2流程：
        1. 从盖章后的 PDF/图片中提取内容
        2. 缩放后插入到目标 Word 文件
        3. 生成最终 Word/PDF
        :param stamp_file: 盖章后的PDF或图片文件
        :param target_word_dir: 目标Word文件目录
        :param image_width: 图片缩放宽度（默认配置中的默认值）
        :return: 生成的Word文件路径列表
        """
        logger.info("开始执行【功能2：盖章页覆盖】")
        try:
            # 1. 获取功能2目录
            images_dir = self.file_manager.get_func2_dir("images")
            result_word_dir = self.file_manager.get_func2_dir("result_word")
            result_pdf_dir = self.file_manager.get_func2_dir("result_pdf")
            target_word_dir = target_word_dir or self.file_manager.get_func2_dir("target_files")

            # 2. 处理盖章文件
            processed_image = self._process_stamp_file(stamp_file, images_dir)

            # 3. 缩放图片（如果需要）
            if image_width:
                scaled_image = images_dir / f"scaled_{processed_image.name}"
                self.image_processor.resize_image(processed_image, scaled_image,
                                                target_width=image_width, keep_ratio=True)
                final_image = scaled_image
            else:
                final_image = processed_image

            # 4. 处理目标Word文件
            word_files = self.file_processor.list_files(target_word_dir, [".docx", ".doc"])
            if not word_files:
                raise BusinessError(f"目标Word目录{target_word_dir}中无Word文件")

            # 5. 批量插入图片
            result_word_files = []
            for word_file in word_files:
                output_word = result_word_dir / f"{word_file.stem}_stamped.docx"

                # 向Word文件插入图片
                self.word_processor.insert_image_to_word(word_file, final_image, output_word)
                result_word_files.append(output_word)

                # 将结果Word转换为PDF
                output_pdf = result_pdf_dir / f"{word_file.stem}_stamped.pdf"
                self.word_processor.word_to_pdf(output_word, output_pdf)
                logger.debug(f"生成最终PDF：{output_pdf}")

            logger.info(f"【功能2】执行完成，成功处理{len(result_word_files)}个Word文件")
            return result_word_files
        except Exception as e:
            logger.error("【功能2】执行失败", exc_info=True)
            raise BusinessError(f"盖章页覆盖失败：{str(e)}") from e

    def _process_stamp_file(self, stamp_file: Path, output_dir: Path) -> Path:
        """
        处理盖章文件：如果是PDF，提取图片；如果是图片，直接复制
        :param stamp_file: 盖章文件路径
        :param output_dir: 处理后文件的输出目录
        :return: 处理后的图片路径
        """
        stamp_suffix = stamp_file.suffix.lower()
        logger.debug(f"处理盖章文件：{stamp_file}，格式：{stamp_suffix}")

        # 验证文件存在
        if not stamp_file.exists() or not stamp_file.is_file():
            raise BusinessError(f"盖章文件不存在：{stamp_file}")

        if stamp_suffix == ".pdf":
            # 从PDF中提取图片
            extracted_images = self.pdf_processor.extract_images(stamp_file, output_dir)
            if not extracted_images:
                raise ImageProcessError(f"无法从PDF中提取图片：{stamp_file}")
            # 返回第一张提取的图片
            return extracted_images[0]
        elif stamp_suffix in [".png", ".jpg", ".jpeg"]:
            # 直接复制图片
            import shutil
            output_image = output_dir / stamp_file.name
            shutil.copy2(stamp_file, output_image)
            logger.debug(f"复制图片：{output_image}")
            return output_image
        else:
            raise BusinessError(f"不支持的盖章文件格式：{stamp_file}")
