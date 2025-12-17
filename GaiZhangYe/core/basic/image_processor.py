# GaiZhangYe/core/image_processor.py
"""
图片处理核心：基于Pillow实现图片相关操作
"""
from PIL import Image
from pathlib import Path
from GaiZhangYe.utils.logger import get_logger
from GaiZhangYe.core.models.exceptions import ImageProcessError

logger = get_logger(__name__)


class ImageProcessor:
    """图片处理器"""

    def resize_image(self, input_image: Path, output_image: Path, target_width: int = None,
                    target_height: int = None, keep_ratio: bool = True) -> None:
        """
        缩放图片
        :param input_image: 输入图片路径
        :param output_image: 输出图片路径
        :param target_width: 目标宽度（像素）
        :param target_height: 目标高度（像素）
        :param keep_ratio: 是否保持宽高比，默认True
        """
        if not input_image.exists():
            raise ImageProcessError(f"图片不存在：{input_image}")
        if not target_width and not target_height:
            raise ImageProcessError("必须提供目标宽度或高度")

        try:
            with Image.open(input_image) as img:
                # 获取原始尺寸
                original_width, original_height = img.size
                logger.debug(f"原始图片尺寸：{original_width}x{original_height}")

                # 计算新尺寸
                if keep_ratio:
                    if target_width and not target_height:
                        # 按目标宽度等比缩放
                        scale_factor = target_width / original_width
                        target_height = int(original_height * scale_factor)
                    elif target_height and not target_width:
                        # 按目标高度等比缩放
                        scale_factor = target_height / original_height
                        target_width = int(original_width * scale_factor)
                # 如果keep_ratio为False且同时提供了宽高，则直接使用

                new_size = (target_width, target_height)
                logger.debug(f"缩放后图片尺寸：{new_size}")

                # 缩放图片（使用ANTIALIAS滤镜保持质量）
                resized_img = img.resize(new_size, Image.LANCZOS)

                # 保存图片
                resized_img.save(output_image)
                logger.info(f"图片缩放成功：{input_image} → {output_image}")
        except Exception as e:
            logger.error(f"图片缩放失败：{input_image}", exc_info=True)
            raise ImageProcessError(f"缩放失败：{str(e)}") from e

    def convert_image_format(self, input_image: Path, output_image: Path, format: str) -> None:
        """
        转换图片格式
        :param input_image: 输入图片路径
        :param output_image: 输出图片路径
        :param format: 目标格式（如'JPEG', 'PNG'）
        """
        if not input_image.exists():
            raise ImageProcessError(f"图片不存在：{input_image}")

        try:
            with Image.open(input_image) as img:
                # 如果是JPEG格式且图片有alpha通道，先转换为RGB
                if format.upper() == "JPEG" and img.mode == "RGBA":
                    img = img.convert("RGB")

                img.save(output_image, format=format.upper())
                logger.info(f"图片格式转换成功：{input_image} → {output_image} ({format})")
        except Exception as e:
            logger.error(f"图片格式转换失败：{input_image}", exc_info=True)
            raise ImageProcessError(f"转换失败：{str(e)}") from e

    def check_image_valid(self, image_path: Path) -> bool:
        """
        检查图片文件是否有效
        :param image_path: 图片路径
        :return: 有效返回True，否则返回False
        """
        if not image_path.exists() or not image_path.is_file():
            return False

        try:
            with Image.open(image_path) as img:
                img.verify()  # 验证图片完整性
            return True
        except Exception:
            return False
