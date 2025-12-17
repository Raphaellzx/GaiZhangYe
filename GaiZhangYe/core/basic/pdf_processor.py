# GaiZhangYe/core/pdf_processor.py
"""
PDF处理核心：基于pymupdf实现PDF相关操作
"""
import fitz  # pymupdf
from pathlib import Path
from typing import List, Tuple
from GaiZhangYe.utils.logger import get_logger
from GaiZhangYe.core.models.exceptions import PdfProcessError

logger = get_logger(__name__)


class PdfProcessor:
    """PDF处理器"""

    def extract_pages(self, source_pdf: Path, target_pages: List[int], output_pdf: Path) -> None:
        """
        从PDF中提取指定页面生成新的PDF
        :param source_pdf: 源PDF文件路径
        :param target_pages: 要提取的页码列表（1-based）
        :param output_pdf: 输出PDF文件路径
        """
        if not source_pdf.exists() or source_pdf.suffix.lower() != ".pdf":
            raise PdfProcessError(f"无效的PDF文件：{source_pdf}")

        try:
            with fitz.open(source_pdf) as doc:
                # 检查页码有效性
                max_pages = doc.page_count
                invalid_pages = [page for page in target_pages if page < 1 or page > max_pages]
                if invalid_pages:
                    raise PdfProcessError(f"无效页码：{invalid_pages}，最大页码：{max_pages}")

                # 创建输出文档
                with fitz.open() as output_doc:
                    # 添加指定页面（注意：fitz是0-based索引）
                    for page_num in target_pages:
                        output_doc.insert_pdf(doc, from_page=page_num - 1, to_page=page_num - 1)

                    # 保存输出
                    output_doc.save(output_pdf)
                    logger.info(f"PDF页面提取成功：{source_pdf} 页码{target_pages} → {output_pdf}")
        except Exception as e:
            logger.error(f"PDF页面提取失败：{source_pdf}", exc_info=True)
            raise PdfProcessError(f"提取失败：{str(e)}") from e

    def extract_images(self, pdf_path: Path, output_dir: Path, dpi: int = 300) -> List[Path]:
        """
        从PDF中提取所有图片
        :param pdf_path: PDF文件路径
        :param output_dir: 图片输出目录
        :param dpi: 图片分辨率，默认300
        :return: 提取的图片路径列表
        """
        if not pdf_path.exists() or pdf_path.suffix.lower() != ".pdf":
            raise PdfProcessError(f"无效的PDF文件：{pdf_path}")

        output_dir.mkdir(exist_ok=True, parents=True)
        extracted_images = []

        try:
            with fitz.open(pdf_path) as doc:
                # 遍历每一页
                for page_num in range(doc.page_count):
                    page = doc[page_num]
                    # 获取页面上的所有图片
                    image_list = page.get_images(full=True)

                    for img_index, img in enumerate(image_list):
                        xref = img[0]
                        base_image = doc.extract_image(xref)
                        image_data = base_image["image"]
                        image_ext = base_image["ext"]

                        # 生成图片文件名
                        image_file = output_dir / f"{pdf_path.stem}_page{page_num + 1}_{img_index + 1}.{image_ext}"

                        # 保存图片
                        with open(image_file, "wb") as f:
                            f.write(image_data)

                        extracted_images.append(image_file)
                        logger.debug(f"从PDF提取图片：{image_file}")

                logger.info(f"从PDF提取图片完成：{pdf_path} → 共{len(extracted_images)}张")
                return extracted_images
        except Exception as e:
            logger.error(f"PDF图片提取失败：{pdf_path}", exc_info=True)
            raise PdfProcessError(f"提取失败：{str(e)}") from e

    def get_page_count(self, pdf_path: Path) -> int:
        """
        获取PDF的总页数
        :param pdf_path: PDF文件路径
        :return: 总页数
        """
        if not pdf_path.exists() or pdf_path.suffix.lower() != ".pdf":
            raise PdfProcessError(f"无效的PDF文件：{pdf_path}")

        try:
            with fitz.open(pdf_path) as doc:
                return doc.page_count
        except Exception as e:
            logger.error(f"获取PDF页数失败：{pdf_path}", exc_info=True)
            raise PdfProcessError(f"获取失败：{str(e)}") from e
