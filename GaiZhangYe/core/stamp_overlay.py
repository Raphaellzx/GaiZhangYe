from pathlib import Path
from typing import List, Optional
import shutil
import os
from GaiZhangYe.utils.logger import get_logger
from GaiZhangYe.core.basic.file_manager import windows_natural_sort_key
from GaiZhangYe.core.basic.file_manager import get_file_manager
from GaiZhangYe.core.basic.word_processor import WordProcessor
from GaiZhangYe.core.basic.pdf_processor import PdfProcessor
from GaiZhangYe.core.basic.image_processor import ImageProcessor
from GaiZhangYe.core.models.exceptions import BusinessError
from GaiZhangYe.core.models.data_models import Func2Data, StampConfig

logger = get_logger(__name__)


class StampOverlayService:
    """盖章页覆盖服务"""

    def __init__(self):
        self.file_manager = get_file_manager()
        self.word_processor = WordProcessor()
        self.pdf_processor = PdfProcessor()
        self.image_processor = ImageProcessor()

    def run(self, target_word_dir: Optional[Path] = None,
            image_width: Optional[int] = None, image_files: Optional[List[Path]] = None,
            configs: Optional[Func2Data] = None,
            result_word_dir: Optional[Path] = None, result_pdf_dir: Optional[Path] = None) -> List[Path]:
        """
        执行功能2流程：
        1. 缩放图片后插入到目标 Word 文件
        2. 生成最终 Word/PDF

        :param target_word_dir: 目标Word文件目录（可选）
        :param image_width: 图片缩放宽度（可选）
        :param image_files: 盖章图片文件列表（可选）
        :param configs: 配置数据模型（可选）
        :param result_word_dir: 输出Word文件目录（可选）
        :param result_pdf_dir: 输出PDF文件目录（可选）
        :return: 生成的Word文件路径列表
        """
        logger.info("开始执行【功能2：盖章页覆盖】")
        try:
            images_dir, final_result_word_dir, final_result_pdf_dir, target_word_dir = self._init_directories(
                target_word_dir, result_word_dir, result_pdf_dir)

            has_valid_config = configs and configs.configs

            if not has_valid_config:
                self._validate_images(image_files)

            word_files = self._get_target_word_files(target_word_dir)
            sorted_word_files = sorted(word_files, key=windows_natural_sort_key)

            sorted_images = []
            if image_files:
                sorted_images = sorted(image_files, key=windows_natural_sort_key)

            result_word_files = self._batch_insert_images_and_convert(
                sorted_word_files, images_dir, final_result_word_dir, final_result_pdf_dir,
                image_width, configs, sorted_images)

            logger.info(f"【功能2】执行完成，成功处理{len(result_word_files)}个Word文件")
            return result_word_files
        except Exception as e:
            logger.error("【功能2】执行失败", exc_info=True)
            raise BusinessError(f"盖章页覆盖失败：{str(e)}") from e

    def _init_directories(self, target_word_dir: Optional[Path],
                          result_word_dir: Optional[Path],
                          result_pdf_dir: Optional[Path]) -> tuple:
        """初始化功能2所需的目录"""
        images_dir = self.file_manager.get_func2_dir("images")

        if target_word_dir:
            self.file_manager.set_custom_func2_dir("target_files", target_word_dir)
        if result_word_dir:
            self.file_manager.set_custom_func2_dir("result_word", Path(result_word_dir))
        if result_pdf_dir:
            self.file_manager.set_custom_func2_dir("result_pdf", Path(result_pdf_dir))

        target_word_dir = self.file_manager.get_func2_dir("target_files")
        final_result_word_dir = self.file_manager.get_func2_dir("result_word")
        final_result_pdf_dir = self.file_manager.get_func2_dir("result_pdf")

        final_result_word_dir.mkdir(parents=True, exist_ok=True)
        final_result_pdf_dir.mkdir(parents=True, exist_ok=True)

        return images_dir, final_result_word_dir, final_result_pdf_dir, target_word_dir

    def _validate_images(self, image_files: Optional[List[Path]]) -> None:
        """验证图片文件列表是否为空以及每个图片文件是否存在"""
        if not image_files or len(image_files) == 0:
            raise BusinessError("没有可处理的图片文件")

        # 验证每个图片文件是否存在
        for img_path in image_files:
            if isinstance(img_path, str):
                img_path = Path(img_path)
            if not img_path.exists():
                raise BusinessError(f"图片文件不存在：{img_path}")

    def _get_target_word_files(self, target_word_dir: Path) -> List[Path]:
        """获取目标Word文件目录中的所有Word文件"""
        word_files = self.file_manager.list_files(target_word_dir, [".docx", ".doc"])
        if not word_files:
            raise BusinessError(f"目标Word目录{target_word_dir}中无Word文件")
        return word_files

    def _get_current_config(self, filename: str, configs: Func2Data) -> Optional[StampConfig]:
        """根据文件名获取对应的配置信息"""
        if not configs or not configs.configs:
            return None

        for config in configs.configs:
            if config.filename == filename:
                return config

        return None

    def _batch_insert_images_and_convert(self, sorted_word_files: List[Path],
                                         images_dir: Path, result_word_dir: Path, result_pdf_dir: Path,
                                         image_width: int, configs: Func2Data, sorted_images: List[Path] = None) -> List[Path]:
        """批量插入图片并将结果转换为PDF"""
        result_word_files = []
        image_index = 0
        total_images = len(sorted_images) if sorted_images else 0

        for word in sorted_word_files:
            output_word = result_word_dir / f"{word.stem}.docx"
            current_config = self._get_current_config(word.name, configs)
            processed_successfully = False

            if current_config:
                processed_successfully = self._process_with_config(
                    current_config, word, output_word, images_dir, image_width)

            if not processed_successfully and sorted_images and image_index < total_images:
                processed_successfully = self._process_with_default_mode(
                    word, output_word, images_dir, image_width, sorted_images, image_index)
                if processed_successfully:
                    image_index += 1

            if processed_successfully:
                result_word_files.append(output_word)
                output_pdf = result_pdf_dir / f"{output_word.stem}.pdf"
                self.word_processor.word_to_pdf(output_word, output_pdf)
            else:
                logger.warning(f"无法处理 Word 文件 {word.name}")

        self._ensure_pdfs_generated(result_word_files, result_pdf_dir)
        return result_word_files

    def _process_with_config(self, current_config: StampConfig, word: Path,
                             output_word: Path, images_dir: Path, image_width: int) -> bool:
        """使用配置模式处理Word文件"""
        logger.info(f"[UI配置模式] 处理 Word 文件 {word.name}")

        if not hasattr(current_config, 'image_files') or not hasattr(current_config, 'positions'):
            logger.warning(f"[UI配置模式] Word 文件 {word.name} 的配置不完整")
            return False
        if not current_config.image_files or not current_config.positions:
            logger.warning(f"[UI配置模式] Word 文件 {word.name} 的配置缺少图片或位置信息")
            return False
        if len(current_config.image_files) != len(current_config.positions):
            logger.warning(f"[UI配置模式] Word 文件 {word.name} 的图片数量和位置数量不一致")
            return False

        # 确保输出目录存在并具有写入权限
        output_word.parent.mkdir(parents=True, exist_ok=True)

        # 创建临时文件（使用安全的方式）
        temp_output = output_word.parent / f"{word.stem}_temp.docx"

        try:
            # 尝试复制文件，如果临时文件已存在则先删除
            if temp_output.exists():
                try:
                    os.unlink(temp_output)
                    logger.debug(f"已删除存在的临时文件: {temp_output}")
                except Exception as e:
                    logger.warning(f"删除临时文件失败: {e}")
                    # 尝试使用不同的临时文件名
                    temp_output = output_word.parent / f"{word.stem}_temp_{os.getpid()}.docx"

            shutil.copy2(word, temp_output)
            logger.debug(f"已创建临时文件: {temp_output}")
        except Exception as e:
            logger.error(f"创建临时文件失败: {e}")
            # 尝试直接使用输出文件作为临时文件（如果无法创建临时文件）
            temp_output = output_word
            if not temp_output.exists():
                shutil.copy2(word, temp_output)

        actual_image_paths = []
        for img_path_str in current_config.image_files:
            img_path = Path(img_path_str)
            if img_path.exists():
                actual_image_path = img_path
            else:
                actual_image_path = images_dir / img_path_str
                if not actual_image_path.exists():
                    logger.warning(f"配置中指定的图片不存在：{img_path_str}")
                    continue
            actual_image_paths.append(str(actual_image_path))

        normalized_positions = self._normalize_positions(
            current_config.positions, word)

        for img_input, position in zip(actual_image_paths, normalized_positions):
            logger.info(f"[UI配置模式] 将图片 {img_input} 插入文件 {word.name} 的页码 {position}")
            img_path = Path(img_input) if Path(img_input).exists() else images_dir / img_input

            if not img_path.exists():
                logger.warning(f"图片文件不存在：{img_path}，跳过该图片")
                continue

            # 直接使用原始图片，不进行提前缩放，保持最高清晰度
            final_image = img_path

            try:
                image_page = int(position)
            except Exception:
                try:
                    image_page = self.word_processor.get_word_page_count(word)
                except Exception:
                    image_page = 1

            self.word_processor.insert_image_to_word(temp_output, final_image, image_page, temp_output)

        # 处理完成后，确保输出文件正确生成
        if temp_output != output_word:
            try:
                if os.path.exists(output_word):
                    os.unlink(output_word)
                os.rename(temp_output, output_word)
                logger.debug(f"已重命名临时文件为输出文件: {output_word}")
            except Exception as e:
                logger.error(f"重命名临时文件失败: {e}")
                # 如果重命名失败，尝试直接复制
                try:
                    shutil.copy2(temp_output, output_word)
                    os.unlink(temp_output)
                except Exception as copy_e:
                    logger.error(f"复制临时文件失败: {copy_e}")
                    return False

        logger.info(f"[UI配置模式] 成功处理 Word 文件 {word.name}，插入图片 {len(actual_image_paths)} 张")
        return True

    def _process_with_default_mode(self, word: Path, output_word: Path,
                                  images_dir: Path, image_width: int,
                                  sorted_images: List[Path], image_index: int) -> bool:
        """使用默认模式处理Word文件"""
        logger.info(f"[默认模式] 处理 Word 文件 {word.name}")

        try:
            default_page_for_word = self.word_processor.get_word_page_count(word)
        except Exception:
            default_page_for_word = 1

        temp_config = type('TempConfig', (), {
            'filename': word.name,
            'image_files': [str(sorted_images[image_index])],
            'positions': [default_page_for_word]
        })()

        logger.info(f"[默认模式] 为 {word.name} 使用默认插入页码: {default_page_for_word}")

        return self._process_with_config(temp_config, word, output_word, images_dir, image_width)

    def _normalize_positions(self, positions, word):
        """规范化插入位置"""
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

            if n < 1:
                n = 1
            if total_pages and n > total_pages:
                n = total_pages
            normalized.append(n)

        return normalized


    def _ensure_pdfs_generated(self, result_word_files: List[Path], result_pdf_dir: Path) -> None:
        """确保所有成功处理的Word文件都已转换为PDF"""
        for word_file in result_word_files:
            if word_file.exists():
                pdf_file = self._find_pdf_file(result_pdf_dir, word_file.stem)
                if not pdf_file or not pdf_file.exists():
                    logger.info(f"【安全检查】重新生成 PDF 文件：{result_pdf_dir / word_file.stem}")
                    output_pdf = result_pdf_dir / f"{word_file.stem}.pdf"
                    self.word_processor.word_to_pdf(word_file, output_pdf)

    def _find_pdf_file(self, result_pdf_dir: Path, stem: str) -> Path:
        """在结果目录中查找与给定word stem对应的PDF文件"""
        candidates = []
        try:
            stamped = result_pdf_dir / f"{stem}_stamped.pdf"
            normal = result_pdf_dir / f"{stem}.pdf"
            if stamped.exists():
                return stamped
            if normal.exists():
                return normal

            for p in result_pdf_dir.iterdir():
                if not p.is_file():
                    continue
                if p.stem == stem or p.stem == f"{stem}_stamped":
                    candidates.append(p)
        except Exception:
            return None

        return candidates[0] if candidates else None