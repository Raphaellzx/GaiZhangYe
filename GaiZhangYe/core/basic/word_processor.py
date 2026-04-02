import win32com.client
import pythoncom

# 导入Word常量
from win32com.client import constants as win32

from pathlib import Path
from typing import List, Union
from GaiZhangYe.core.models.data_models import ImageLocationType
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
        try:
            # 接受所有修订
            if doc.Revisions.Count > 0:
                doc.Revisions.AcceptAll()
                logger.debug(f"已接受文档{doc.Name}的所有修订")
        except Exception as e:
            logger.warning(f"无法接受文档{doc.Name}的修订：{str(e)}")

        try:
            # 删除所有注释
            if hasattr(doc, 'Comments') and doc.Comments.Count > 0:
                doc.Comments.DeleteAll()
                logger.debug(f"已删除文档{doc.Name}的所有注释")
        except Exception as e:
            logger.warning(f"无法删除文档{doc.Name}的注释：{str(e)}")

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

    def _get_image_dimensions(self, image_path: Path) -> tuple:
        """获取图片的原始尺寸（宽度和高度，单位为像素）"""
        try:
            from PIL import Image
            with Image.open(image_path) as img:
                return img.size
        except Exception as e:
            logger.warning(f"无法获取图片尺寸：{image_path}，使用默认尺寸", exc_info=True)
            return (1000, 1000)  # 默认尺寸

    def __del__(self):
        """析构函数：自动关闭Word应用"""
        self.close()

    def get_word_page_count(self, word_path: Path) -> int:
        """获取Word文件的页数"""
        if word_path.name.startswith("~$"):
            raise WordProcessError(f"临时文件，跳过获取页数：{word_path}")
        if not word_path.exists():
            raise FileNotFoundError(f"Word文件不存在：{word_path}")
        if word_path.suffix.lower() not in [".docx", ".doc"]:
            raise WordProcessError(f"非Word文件：{word_path}")

        try:
            word_app = self._get_word_app()
            doc = word_app.Documents.Open(str(word_path.absolute()), ReadOnly=True)

            try:
                self._clean_doc(doc)
            except Exception:
                pass

            page_count = doc.ComputeStatistics(2)

            # 确保返回的是可序列化的数据类型
            page_count = int(page_count) if page_count else 0

            doc.Close(SaveChanges=False)
            logger.info(f"获取Word文件页数成功：{word_path} - {page_count}页")
            return page_count
        except Exception as e:
            logger.error(f"获取Word文件页数失败：{word_path}", exc_info=True)
            raise WordProcessError(f"获取页数失败：{str(e)}") from e

    def insert_image_to_word(self, word_path: Path, image_path: Path,
                        image_location: Union[int, ImageLocationType],
                        output_path: Path) -> None:
        """向Word插入图片（盖章页覆盖核心）
        :param image_location: 图片插入位置，支持数值页码，或ImageLocationType枚举
        """
        if not word_path.exists():
            raise FileNotFoundError(f"Word文件不存在：{word_path}")
        if not image_path.exists():
            raise FileNotFoundError(f"图片文件不存在：{image_path}")

        doc = None
        try:
            word_app = self._get_word_app()
            doc = word_app.Documents.Open(str(word_path.absolute()))
            doc.Activate()

            # 确定插入页码
            if isinstance(image_location, ImageLocationType):
                if image_location == ImageLocationType.LAST_PAGE:
                    page_num = doc.ComputeStatistics(2)  # 总页数
                else:
                    page_num = 1  # FIRST_PAGE
            else:
                try:
                    page_num = int(image_location)
                    max_pages = doc.ComputeStatistics(2)
                    page_num = max(1, min(page_num, max_pages))
                except Exception:
                    raise WordProcessError(f"无效的图片位置：{image_location}")

            logger.info(f"图片将插入的目标页: {page_num}")

            # 定位到目标页的起始位置（使用更可靠的方法）
            rng = None
            try:
                # 先选择整个文档内容
                rng = doc.Content

                # 定位到第一个字符
                rng.Collapse(Direction=0)  # wdCollapseStart = 0

                # 逐页移动到目标页
                for i in range(1, page_num):
                    rng.MoveStart(Unit=2, Count=1)  # wdCharacter = 1, wdPage = 2

                logger.debug(f"成功定位到页面 {page_num}")
            except Exception as e:
                logger.warning(f"定位到页面 {page_num} 失败，使用回退方案: {e}")
                try:
                    rng = doc.GoTo(What=1, Which=1, Count=page_num)  # wdGoToPage = 1
                except Exception:
                    try:
                        rng = doc.GoTo(1, 1, page_num)
                    except Exception:
                        rng = doc.Content
                        try:
                            rng.Collapse(Direction=0)
                        except Exception:
                            pass

            # 获取Range对象
            try:
                test_range = getattr(rng, 'Range', None)
                if test_range is not None:
                    rng = test_range
            except Exception:
                pass

            # 插入图片为InlineShape
            inline_shapes = rng.InlineShapes
            inline_shape = inline_shapes.AddPicture(str(image_path.absolute()))

            # 转换为浮动图形并铺满整页
            try:
                shp = inline_shape.ConvertToShape()

                # 获取目标页的section
                current_section = None
                try:
                    current_section = rng.Sections(1)
                except Exception:
                    try:
                        current_section = doc.Sections(doc.Sections.Count)
                    except Exception:
                        current_section = None

                if current_section is not None:
                    page_height = current_section.PageSetup.PageHeight
                    page_width = current_section.PageSetup.PageWidth

                    # 设置图片格式（浮于文字上方并铺满整页）
                    shp.WrapFormat.Type = 3  # 浮于文字上方
                    shp.RelativeHorizontalPosition = 1  # 相对于页面边缘
                    shp.RelativeVerticalPosition = 1    # 相对于页面边缘

                    # 根据页面方向设置图片尺寸和旋转
                    if page_height > page_width:
                        # 纵向页面
                        shp.Width = page_width
                        shp.Height = page_height
                    else:
                        # 横向页面，旋转90度
                        shp.Width = page_height
                        shp.Height = page_width
                        try:
                            shp.Rotation = 90
                        except Exception:
                            pass

                    # 使用负值确保覆盖整个页面（不受边距影响）
                    try:
                        shp.Left = -999995
                        shp.Top = -999995
                    except Exception:
                        pass

                    # 置于顶层
                    try:
                        shp.ZOrder(1)  # wdBringToFront = 1
                    except Exception:
                        pass
                else:
                    # 无法获取页面设置时的回退方案
                    if inline_shape.Width > 400:
                        ratio = 400 / inline_shape.Width
                        inline_shape.Width = 400
                        inline_shape.Height = int(inline_shape.Height * ratio)

            except Exception as e:
                # ConvertToShape失败时的回退方案
                logger.warning(f"转换为浮动图形失败，使用inline方式：{e}")
                try:
                    current_section = None
                    try:
                        current_section = rng.Sections(1)
                    except Exception:
                        try:
                            current_section = doc.Sections(doc.Sections.Count)
                        except Exception:
                            current_section = None

                    if current_section is not None:
                        page_height = current_section.PageSetup.PageHeight
                        page_width = current_section.PageSetup.PageWidth
                        left_margin = current_section.PageSetup.LeftMargin
                        right_margin = current_section.PageSetup.RightMargin
                        usable_width = max(1, page_width - left_margin - right_margin)
                        
                        if inline_shape.Width > usable_width:
                            ratio = usable_width / inline_shape.Width
                            inline_shape.Width = usable_width
                            inline_shape.Height = int(inline_shape.Height * ratio)
                except Exception:
                    pass

            # 保存文档
            doc.SaveAs(str(output_path.absolute()))
            logger.info(f"图片插入Word成功：{image_path} → {output_path}")

        except Exception as e:
            logger.error(f"插入图片失败：{word_path}", exc_info=True)
            raise WordProcessError(f"插入失败：{str(e)}") from e
        finally:
            if doc:
                try:
                    doc.Close(SaveChanges=False)
                except Exception as e:
                    logger.warning(f"关闭文档失败：{str(e)}")