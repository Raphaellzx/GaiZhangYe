# GaiZhangYe/core/word_processor.py
import win32com.client
import pythoncom

# 动态生成常量
win32 = win32com.client.constants
from pathlib import Path
from typing import List

from GaiZhangYe.utils.logger import get_logger
from GaiZhangYe.core.models.exceptions import WordProcessError

logger = get_logger(__name__)

class WordProcessor:
    """Word处理器：封装pywin32的Word操作"""
    def __init__(self):
        self._word_app = None

    def _get_word_app(self) -> win32com.client.CDispatch:
        """获取Word应用实例（单例）"""
        pythoncom.CoInitialize()
        if not self._word_app:
            # 使用DispatchEx避免冲突，后台运行
            self._word_app = win32com.client.DispatchEx("Word.Application")
            self._word_app.Visible = False
            self._word_app.DisplayAlerts = 0  # 抑制弹窗
        return self._word_app

    def _clean_doc(self, doc):
        """清理文档：接受所有修订 + 删除所有注释"""
        # 接受所有修订
        if doc.Revisions.Count > 0:
            doc.Revisions.AcceptAll()
            logger.debug(f"已接受文档{doc.Name}的所有修订")
        # 删除所有注释
        if doc.Comments.Count > 0:
            doc.Comments.DeleteAll()
            logger.debug(f"已删除文档{doc.Name}的所有注释")

    def word_to_pdf(self, word_path: Path, pdf_path: Path) -> None:
        """单文件Word转PDF（含修订/注释清理）"""
        # 前置校验
        if not word_path.exists():
            raise FileNotFoundError(f"Word文件不存在：{word_path}")
        if word_path.suffix.lower() not in [".docx", ".doc"]:
            raise WordProcessError(f"非Word文件（仅支持.doc/.docx）：{word_path}")
        
        doc = None
        try:
            word_app = self._get_word_app()
            # 打开文档（绝对路径避免解析问题）
            doc = word_app.Documents.Open(str(word_path.absolute()))
            
            # 清理修订和注释
            self._clean_doc(doc)

            # 确保输出目录存在
            pdf_path.parent.mkdir(parents=True, exist_ok=True)
            
            # 导出PDF（使用直接常量值以避免AttributeError）
            doc.ExportAsFixedFormat(
                OutputFileName=str(pdf_path.absolute()),
                ExportFormat=17,  # wdExportFormatPDF = 17
                OpenAfterExport=False,  # 导出后不打开PDF
                Item=0,  # wdExportDocument = 0 (仅导出正文)
                CreateBookmarks=1  # wdExportCreateHeadingBookmarks = 1
            )
            
            logger.info(f"Word转PDF成功：{word_path} → {pdf_path}")
        except Exception as e:
            logger.error(f"Word转PDF失败：{word_path}", exc_info=True)
            raise WordProcessError(f"转换失败：{str(e)}") from e
        finally:
            # 确保文档关闭，释放资源
            if doc:
                doc.Close(SaveChanges=False)  # 不保存原文档的修改

    def batch_word_to_pdf(self, input_dir: Path, output_dir: Path) -> List[Path]:
        """批量Word转PDF"""
        # 校验输入目录
        if not input_dir.exists():
            raise WordProcessError(f"输入目录不存在：{input_dir}")
        
        # 获取所有Word文件（去重，Windows系统大小写不敏感）
        word_files = set()
        for ext in ["*.docx", "*.doc", "*.DOCX", "*.DOC"]:
            word_files.update(input_dir.glob(ext))

        word_files = list(word_files)
        if not word_files:
            raise WordProcessError(f"目录{input_dir}无Word文件（.doc/.docx）")

        pdf_paths = []
        for word_file in word_files:
            try:
                # 构造PDF输出路径
                pdf_path = output_dir / f"{word_file.stem}.pdf"
                # 调用单文件转换（已包含修订/注释清理）
                self.word_to_pdf(word_file, pdf_path)
                pdf_paths.append(pdf_path)
            except (FileNotFoundError, WordProcessError) as e:
                # 单个文件失败不中断批量流程，仅记录日志
                logger.warning(f"跳过文件{word_file}：{str(e)}")
                continue

        logger.info(f"批量转换完成，成功生成{len(pdf_paths)}/{len(word_files)}个PDF文件")
        return pdf_paths

    def close(self):
        """手动关闭Word应用，释放资源"""
        if self._word_app:
            try:
                self._word_app.Quit()
                logger.info("Word应用已退出")
            except Exception as e:
                logger.warning(f"退出Word应用失败：{str(e)}")
            finally:
                self._word_app = None

    def __del__(self):
        """析构函数：自动关闭Word应用"""
        self.close()

    def get_word_page_count(self, word_path: Path) -> int:
        """获取Word文件的页数"""
        if not word_path.exists():
            raise FileNotFoundError(f"Word文件不存在：{word_path}")
        if word_path.suffix not in [".docx", ".doc"]:
            raise WordProcessError(f"非Word文件：{word_path}")

        try:
            word_app = self._get_word_app()
            # 使用完整的绝对路径字符串，避免编码问题
            abs_path = str(word_path.absolute())
            # 尝试使用不同的参数打开文件
            doc = word_app.Documents.Open(
                FileName=abs_path,
                ReadOnly=True,
                Encoding=1252,  # windows-1252编码
                Revert=True
            )
            page_count = doc.ComputeStatistics(2)  # 2=wdStatisticPages
            doc.Close(SaveChanges=False)  # 只读模式打开，不保存变化
            logger.info(f"获取Word文件页数成功：{word_path} - {page_count}页")
            return page_count
        except Exception as e:
            logger.error(f"获取Word文件页数失败：{word_path}", exc_info=True)
            # 尝试另一种方式：先将文件转换为PDF，然后通过PDF获取页数
            try:
                from pathlib import Path
                import tempfile
                import pymupdf as fitz  # PyMuPDF

                # 创建临时PDF文件
                with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
                    temp_pdf_path = tmp.name

                # 转换为PDF
                self.word_to_pdf(word_path, Path(temp_pdf_path))

                # 从PDF获取页数
                with fitz.open(temp_pdf_path) as doc:
                    page_count = doc.page_count

                # 删除临时文件
                Path(temp_pdf_path).unlink()

                logger.info(f"通过PDF转换获取Word文件页数成功：{word_path} - {page_count}页")
                return page_count
            except Exception as e2:
                logger.error(f"通过PDF转换获取页数也失败：{word_path}", exc_info=True)
                raise WordProcessError(f"获取页数失败：{str(e)}") from e

    def insert_image_to_word(self, word_path: Path, image_path: Path, image_location: str, output_path: Path) -> None:
        """向Word插入图片（盖章页覆盖核心）
        :param image_location: 图片插入位置，可以是页码（数字字符串，如 '1'、'5'）或特殊值 'last_page'（最后一页）
        """
        if not word_path.exists():
            raise FileNotFoundError(f"Word文件不存在：{word_path}")
        if not image_path.exists():
            raise FileNotFoundError(f"图片文件不存在：{image_path}")

        try:
            # 获取Word应用实例
            if not self._word_app:
                self._word_app = self._get_word_app()

            doc = self._word_app.Documents.Open(str(word_path))
            doc.Activate()
            selection = self._word_app.Selection

            # 根据image_location参数定位插入位置
            if image_location in ['last_page', 'last'] or not image_location:
                # 定位到最后一页（物理页面末尾）
                # 2=wdStatisticPages - 计算文档总页数
                target_page = self._word_app.ActiveDocument.ComputeStatistics(2)
                # wdGoToPage=1，wdGoToAbsolute=1
                selection.GoTo(What=1, Which=1, Count=target_page)

            elif image_location.isdigit():
                # 定位到指定页码
                page_num = int(image_location)
                # 转到指定页
                self._word_app.Selection.GoTo(
                    What=1,  # 1=wdGoToPage
                    Which=1,  # 1=wdGoToAbsolute
                    Count=page_num
                )

            elif image_location.startswith('specific_page:'):
                # 解析格式 'specific_page:<页码>'
                page_num = int(image_location.split(':')[1])
                # 转到指定页
                self._word_app.Selection.GoTo(
                    What=1,  # 1=wdGoToPage
                    Which=1,  # 1=wdGoToAbsolute
                    Count=page_num
                )

            # 在当前位置插入图片
            shape = selection.InlineShapes.AddPicture(str(image_path)).ConvertToShape()

            # 设置图片格式
            shape.WrapFormat.Type = 3
            shape.RelativeHorizontalPosition = 1
            shape.RelativeVerticalPosition = 1
            current_section = doc.ActiveWindow.Selection.Sections(1)

            page_height = current_section.PageSetup.PageHeight
            page_width = current_section.PageSetup.PageWidth

            if(page_height>page_width):
                shape.Width = page_width
                shape.Height = page_height
            else:
                shape.Width = page_height
                shape.Height = page_width
                shape.rotation=90

            shape.Left = -999995
            shape.Top = -999995

            doc.SaveAs(str(output_path))
            doc.Close()
            logger.info(f"图片插入Word成功：{image_path} → {output_path}")
        except Exception as e:
            logger.error(f"插入图片失败：{word_path}", exc_info=True)
            raise WordProcessError(f"插入失败：{str(e)}") from e

