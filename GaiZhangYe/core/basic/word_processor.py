# GaiZhangYe/core/word_processor.py
import win32com.client
import pythoncom
from pathlib import Path
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
            self._word_app = win32com.client.DispatchEx("Word.Application")
            self._word_app.Visible = False
            self._word_app.DisplayAlerts = 0  # 抑制弹窗
        return self._word_app

    def word_to_pdf(self, word_path: Path, pdf_path: Path) -> None:
        """单文件Word转PDF"""
        if not word_path.exists():
            raise FileNotFoundError(f"Word文件不存在：{word_path}")
        if word_path.suffix not in [".docx", ".doc"]:
            raise WordProcessError(f"非Word文件：{word_path}")

        try:
            word_app = self._get_word_app()
            doc = word_app.Documents.Open(str(word_path))
            doc.SaveAs(str(pdf_path), FileFormat=17)  # 17=PDF格式
            doc.Close()
            logger.info(f"Word转PDF成功：{word_path} → {pdf_path}")
        except Exception as e:
            logger.error(f"Word转PDF失败：{word_path}", exc_info=True)
            raise WordProcessError(f"转换失败：{str(e)}") from e

    def batch_word_to_pdf(self, input_dir: Path, output_dir: Path) -> list[Path]:
        """批量Word转PDF"""
        word_files = list(input_dir.glob("*.docx")) + list(input_dir.glob("*.doc"))
        if not word_files:
            raise WordProcessError(f"目录{input_dir}无Word文件")

        pdf_paths = []
        for word_file in word_files:
            pdf_path = output_dir / f"{word_file.stem}.pdf"
            self.word_to_pdf(word_file, pdf_path)
            pdf_paths.append(pdf_path)
        return pdf_paths

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
                import fitz  # PyMuPDF

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
            shape.Width = current_section.PageSetup.PageWidth
            shape.Height = current_section.PageSetup.PageHeight
            shape.Left = -999995
            shape.Top = -999995

            doc.SaveAs(str(output_path))
            doc.Close()
            logger.info(f"图片插入Word成功：{image_path} → {output_path}")
        except Exception as e:
            logger.error(f"插入图片失败：{word_path}", exc_info=True)
            raise WordProcessError(f"插入失败：{str(e)}") from e

    def __del__(self):
        """销毁Word实例"""
        if self._word_app:
            self._word_app.Quit()
            pythoncom.CoUninitialize()