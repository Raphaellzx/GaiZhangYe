# GaiZhangYe/core/services/stamp_prepare.py
from pathlib import Path
from typing import Optional, List
from GaiZhangYe.utils.logger import get_logger
from GaiZhangYe.core.basic.file_manager import FileManager
from GaiZhangYe.core.basic.word_processor import WordProcessor
from GaiZhangYe.core.basic.pdf_processor import PdfProcessor
from GaiZhangYe.core.basic.file_processor import FileProcessor
from GaiZhangYe.core.models.exceptions import BusinessError
from GaiZhangYe.core.data_communication import get_data_service

import fitz # PyMuPDF
import os

logger = get_logger(__name__)

class StampPrepareService:
    """功能1：准备盖章页（整合core所有相关模块）"""
    def __init__(self):
        self.file_manager = FileManager()
        self.word_processor = WordProcessor()
        self.pdf_processor = PdfProcessor()
        self.file_processor = FileProcessor()

    def run(self, target_pages: Optional[dict[str, list[int]]] = None, word_dir: Optional[Path] = None) -> list[Path]:
        # 如果没有传入target_pages，从数据文件中读取
        if target_pages is None:
            target_pages = get_data_service().get_func1_data()
        # 如果数据文件中也没有数据，抛出异常
        if not target_pages:
            raise BusinessError("没有找到要提取的页面配置数据")
        """
        :param target_pages: 要提取的页面字典，键为Word文件名，值为页面列表
        """
        """
        执行功能1流程：
        1. 读取Nostamped_Word目录的Word文件 → 转PDF到Nostamped_PDF
        2. 从PDF中提取指定页面 → 保存到Stamped_Pages
        """
        logger.info("开始执行【功能1：准备盖章页】")
        try:
            # 1. 获取功能1目录（优先用传入的word_dir，否则用默认目录）
            nostamped_word_dir = word_dir or self.file_manager.get_func1_dir("nostamped_word")
            nostamped_pdf_dir = self.file_manager.get_func1_dir("nostamped_pdf")
            stamped_pages_dir = self.file_manager.get_func1_dir("stamped_pages")

            # 获取Temp目录，如果不存在则创建
            temp_dir = self.file_manager.get_func1_dir("temp")
            temp_dir.mkdir(exist_ok=True)

            # 2. 校验输入目录有Word文件
            word_files = self.file_processor.list_files(nostamped_word_dir, [".docx", ".doc"])
            if not word_files:
                raise BusinessError(f"目录{nostamped_word_dir}无Word文件")

            # 3. 为每个Word文件直接转换指定页面为PDF，并合并为一个文件
            merged_pdf = fitz.open()

            # 按Word文件名排序
            sorted_word_files = sorted(word_files, key=lambda x: x.name)

            for word_file in sorted_word_files:
                word_filename = word_file.stem
                if word_filename in target_pages:
                    pages_to_extract = target_pages[word_filename]
                    logger.info(f"正在将 {word_file.name} 的页面 {pages_to_extract} 转换为PDF")

                    # 创建临时文件路径 - 保存到Temp目录
                    temp_pdf = temp_dir / f"{word_filename}_temp.pdf"

                    try:
                        # 只转换指定页面
                        # 这里先转换整个Word到PDF，然后提取页面
                        # 注意：win32com的Word API没有直接转换指定页面的功能
                        self.word_processor.word_to_pdf(word_file, temp_pdf)

                        # 打开临时PDF文件并提取指定页面
                        with fitz.open(temp_pdf) as doc:
                            for page_num in pages_to_extract:
                                if 1 <= page_num <= len(doc):
                                    merged_pdf.insert_pdf(doc, from_page=page_num - 1, to_page=page_num - 1)

                                    # 保存提取的页面到nostamped_PDF目录
                                    pdf_output = nostamped_pdf_dir / f"{word_filename}_第{page_num}页.pdf"
                                    with fitz.open() as new_doc:
                                        new_doc.insert_pdf(doc, from_page=page_num - 1, to_page=page_num - 1)
                                        new_doc.save(pdf_output)
                                        logger.info(f"已保存提取的页面：{pdf_output}")

                                else:
                                    logger.warning(f"页码 {page_num} 超出 {word_file.name} 的范围")
                    finally:
                        # 删除临时PDF文件
                        if temp_pdf.exists():
                            try:
                                temp_pdf.unlink()
                                logger.info(f"已删除临时PDF文件：{temp_pdf}")
                            except Exception as e:
                                logger.error(f"删除临时PDF文件失败：{temp_pdf}", exc_info=True)

            # 保存合并后的PDF文件
            if merged_pdf.page_count > 0:
                output_pdf = stamped_pages_dir / "stamped_pages.pdf"
                merged_pdf.save(output_pdf)
                merged_pdf.close()

                logger.info(f"【功能1】执行完成，已将所有指定页面合并为一个PDF文件：{output_pdf}")
                return [output_pdf]
            else:
                merged_pdf.close()
                logger.warning("【功能1】执行完成，但没有转换到任何页面")
                return []
        except Exception as e:
            logger.error("【功能1】执行失败", exc_info=True)
            raise BusinessError(f"准备盖章页失败：{str(e)}") from e