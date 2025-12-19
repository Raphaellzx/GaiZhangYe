# GaiZhangYe/core/services/batch_convert.py
"""
功能3：批量Word转PDF服务
"""
from pathlib import Path
from typing import List
from GaiZhangYe.utils.logger import get_logger
from GaiZhangYe.core.basic.file_processor import FileProcessor
from GaiZhangYe.core.basic.word_processor import WordProcessor
from GaiZhangYe.core.models.exceptions import BusinessError

logger = get_logger(__name__)


class BatchConvertService:
    """批量Word转PDF服务"""

    def __init__(self):
        self.file_processor = FileProcessor()
        self.word_processor = WordProcessor()

    def run(self, input_dir: Path, output_dir: Path) -> List[Path]:
        """
        执行批量Word转PDF
        :param input_dir: 输入目录（包含待转换的Word文件）
        :param output_dir: 输出目录（存放生成的PDF文件）
        :return: 生成的PDF文件路径列表
        """
        logger.info("开始执行【功能3：批量Word转PDF】")
        try:
            # 确保输出目录存在
            output_dir.mkdir(exist_ok=True, parents=True)

            # 1. 列出输入目录中的所有Word文件
            word_files = self.file_processor.list_files(input_dir, [".docx", ".doc"])
            if not word_files:
                raise BusinessError(f"目录{input_dir}中未找到Word文件")

            logger.info(f"找到{len(word_files)}个Word文件待转换")

            # 2. 批量转换
            converted_pdfs = self.word_processor.batch_word_to_pdf(
                input_dir, output_dir
            )

            logger.info(
                f"【功能3】批量转换完成，成功生成{len(converted_pdfs)}个PDF文件"
            )
            return converted_pdfs
        except Exception as e:
            logger.error("【功能3】批量转换失败", exc_info=True)
            raise BusinessError(f"批量转PDF失败：{str(e)}") from e
