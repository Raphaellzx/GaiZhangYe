# GaiZhangYe/core/services/stamp_overlay.py
"""
功能2：盖章页覆盖服务
"""
import re
from pathlib import Path
from typing import List
from GaiZhangYe.utils.logger import get_logger
from GaiZhangYe.core.basic.file_manager import FileManager
from GaiZhangYe.core.basic.file_processor import FileProcessor
from GaiZhangYe.core.basic.word_processor import WordProcessor
from GaiZhangYe.core.basic.pdf_processor import PdfProcessor
from GaiZhangYe.core.basic.image_processor import ImageProcessor
from GaiZhangYe.core.models.exceptions import BusinessError
from GaiZhangYe.core.data_communication import get_data_service

logger = get_logger(__name__)


def windows_natural_sort_key(filename):
    """与FuGai.py相同的自然排序实现"""
    import re

    s = filename.name
    return [int(text) if text.isdigit() else text.lower()
            for text in re.split(r'(\d+)', s)]


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
            # 如果没有提供configs，才需要验证image_files
            # configs 可能是 None 或者 空字典 {}
            if configs is None or (isinstance(configs, dict) and len(configs) == 0):
                # 确保在没有配置时，image_files 非空且已验证
                self._validate_images(image_files)

            # 另外，如果有配置但image_files是None，不需要验证
            # 因为配置模式会从配置中获取图片

            # 3. 获取并处理目标Word文件
            word_files = self._get_target_word_files(target_word_dir)

            # 应用自然排序
            sorted_word_files = sorted(word_files, key=windows_natural_sort_key)

            # 根据情况生成sorted_images
            sorted_images = []
            if not configs:
                # 默认模式下需要sorted_images
                sorted_images = sorted(image_files, key=windows_natural_sort_key)

            # 5. 批量插入图片并转换为PDF
            result_word_files = self._batch_insert_images_and_convert(
                sorted_word_files, images_dir, result_word_dir, result_pdf_dir, image_width, configs, sorted_images)

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

                    temp_config = type('TempConfig', (), {
                        'filename': word.name,
                        'image_files': [str(sorted_images[image_index])],
                        'insert_positions': ['last_page']  # 默认插入到最后一页
                    })()

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

                        logger.info(f"[默认模式] 成功处理 Word 文件 {word.name}，插入图片 1 张")
                        processed_successfully = True
                except Exception as e:
                    logger.error(f"[默认模式] 处理失败 Word 文件 {word.name}：{str(e)}")

            if not processed_successfully:
                if current_config:
                    logger.warning(f"已无足够图片或配置错误，无法处理 Word 文件 {word.name}")
                else:
                    logger.warning(f"已无足够图片，无法处理 Word 文件 {word.name}")

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
            return Config(
                filename=filename,
                image_files=[item['image'] for item in items],
                insert_positions=[item['position'] for item in items]
            )
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
            # 只支持数字格式的页码
            if position:  # 非空值
                image_location = str(position)  # 转换为字符串格式
            else:
                image_location = "1"  # 默认插入到第一页

            # 插入图片
            self.word_processor.insert_image_to_word(temp_output, final_image, image_location, temp_output)

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

        # 插入图片
        self.word_processor.insert_image_to_word(word, final_image, "last_page", output_word)
        return True

    def _convert_word_to_pdf(self, output_word: Path, result_pdf_dir: Path) -> None:
        """将Word文件转换为PDF"""
        output_pdf = result_pdf_dir / f"{output_word.stem}_stamped.pdf"
        self.word_processor.word_to_pdf(output_word, output_pdf)
        logger.debug(f"生成最终PDF：{output_pdf}")