# GaiZhangYe/core/services/stamp_prepare.py
from pathlib import Path
from GaiZhangYe.utils.logger import get_logger
from GaiZhangYe.core.file_manager import FileManager
from GaiZhangYe.core.word_processor import WordProcessor
from GaiZhangYe.core.pdf_processor import PdfProcessor
from GaiZhangYe.core.file_processor import FileProcessor
from GaiZhangYe.core.models.exceptions import BusinessError

logger = get_logger(__name__)

class StampPrepareService:
    """功能1：准备盖章页（整合core所有相关模块）"""
    def __init__(self):
        self.file_manager = FileManager()
        self.word_processor = WordProcessor()
        self.pdf_processor = PdfProcessor()
        self.file_processor = FileProcessor()

    def run(self, target_pages: list[int], word_dir: Optional[Path] = None) -> Path:
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

            # 2. 校验输入目录有Word文件
            word_files = self.file_processor.list_files(nostamped_word_dir, [".docx", ".doc"])
            if not word_files:
                raise BusinessError(f"目录{nostamped_word_dir}无Word文件")

            # 3. Word批量转PDF
            self.word_processor.batch_word_to_pdf(nostamped_word_dir, nostamped_pdf_dir)

            # 4. 取第一个PDF提取指定页面（可扩展为多选）
            pdf_files = self.file_processor.list_files(nostamped_pdf_dir, [".pdf"])
            source_pdf = pdf_files[0]
            output_pdf = stamped_pages_dir / f"{source_pdf.stem}_stamped_pages.pdf"
            self.pdf_processor.extract_pages(source_pdf, target_pages, output_pdf)

            logger.info(f"【功能1】执行完成，盖章页PDF：{output_pdf}")
            return output_pdf
        except Exception as e:
            logger.error("【功能1】执行失败", exc_info=True)
            raise BusinessError(f"准备盖章页失败：{str(e)}") from e