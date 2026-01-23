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

    def __del__(self):
        """析构函数：自动关闭Word应用"""
        self.close()

    def get_word_page_count(self, word_path: Path) -> int:
        """获取Word文件的页数"""
        # 跳过Word的临时文件（以~$开头）
        if word_path.name.startswith("~$"):
            raise WordProcessError(f"临时文件，跳过获取页数：{word_path}")

        if not word_path.exists():
            raise FileNotFoundError(f"Word文件不存在：{word_path}")

        # 后缀比较使用不区分大小写的方式
        if word_path.suffix.lower() not in [".docx", ".doc"]:
            raise WordProcessError(f"非Word文件：{word_path}")

        try:
            word_app = self._get_word_app()
            abs_path = str(word_path.absolute())

            # 以只读方式打开，避免弹窗和写锁
            try:
                doc = word_app.Documents.Open(FileName=abs_path, ReadOnly=True)
            except Exception:
                # 有些Word在打开时对参数敏感，重试更简单的调用
                doc = word_app.Documents.Open(abs_path)

            # 清理修订与注释，保证页数统计一致
            try:
                self._clean_doc(doc)
            except Exception:
                pass

            # 尝试强制重排页面并更新域（在复杂文档中有时需要）
            try:
                try:
                    doc.Repaginate()
                except Exception:
                    # 有的Word版本不提供 Repaginate，忽略
                    pass
                try:
                    doc.Fields.Update()
                except Exception:
                    pass
            except Exception:
                pass

            page_count = doc.ComputeStatistics(2)  # 2=wdStatisticPages
            # 在极少数情况下返回0，尝试再次重排
            if not page_count:
                try:
                    doc.Repaginate()
                    page_count = doc.ComputeStatistics(2)
                except Exception:
                    pass

            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass

            logger.info(f"获取Word文件页数成功：{word_path} - {page_count}页")
            return int(page_count)
        except Exception as e:
            logger.error(f"获取Word文件页数失败：{word_path}", exc_info=True)
            # 尝试另一种方式：先将文件转换为PDF，然后通过PDF获取页数
            try:
                import tempfile
                import pymupdf as fitz  # PyMuPDF

                # 创建临时PDF文件
                with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
                    temp_pdf_path = tmp.name

                # 转换为PDF（word_to_pdf会检查原文件存在性）
                self.word_to_pdf(word_path, Path(temp_pdf_path))

                # 从PDF获取页数
                with fitz.open(temp_pdf_path) as pdf_doc:
                    page_count = pdf_doc.page_count

                # 删除临时文件
                try:
                    Path(temp_pdf_path).unlink()
                except Exception:
                    pass

                logger.info(f"通过PDF转换获取Word文件页数成功：{word_path} - {page_count}页")
                return int(page_count)
            except Exception:
                logger.error(f"通过PDF转换获取页数也失败：{word_path}", exc_info=True)
                raise WordProcessError(f"获取页数失败：{str(e)}") from e

    def insert_image_to_word(self, word_path: Path, image_path: Path, image_location, output_path: Path) -> None:
        """向Word插入图片（盖章页覆盖核心）
        :param image_location: 图片插入位置，支持数值页码，或字符串 'last'/'end' 表示最后一页
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

            # image_location expected to be an integer page number provided by frontend
            try:
                page_num = int(image_location)
            except Exception:
                raise WordProcessError(f"image_location must be an integer page number, got: {image_location}")

            # 计算目标页：根据文档页数进行边界修正
            try:
                max_pages = doc.ComputeStatistics(2)
            except Exception:
                max_pages = None

            if max_pages is not None:
                final_target_page = max(1, min(page_num, max_pages))
            else:
                final_target_page = max(1, page_num)

            logger.info(f"图片将插入的目标页: {final_target_page}")

            # 使用 Range/GoTo 来定位到目标页的开头（更稳定，避免依赖 ActiveWindow.Selection）
            rng = None
            try:
                if final_target_page is not None:
                    try:
                        rng = doc.GoTo(What=win32.wdGoToPage, Which=win32.wdGoToAbsolute, Count=final_target_page)
                    except Exception:
                        try:
                            rng = doc.GoTo(1, 1, final_target_page)
                        except Exception:
                            rng = None
                if rng is None:
                    # 回退到文档末尾作为锚点
                    rng = doc.Content
                    try:
                        rng.Collapse(Direction=win32.wdCollapseEnd)
                    except Exception:
                        pass
            except Exception as e:
                logger.warning(f"基于 Range 的定位失败，将使用文档末尾作为插入点：{e}")
                rng = doc.Content
                try:
                    rng.Collapse(Direction=win32.wdCollapseEnd)
                except Exception:
                    pass

            # 在定位 Range 上插入图片为 InlineShape（直接使用 Range 的 InlineShapes）
            try:
                # 如果GoTo返回的是Selection对象，尝试从中获取Range
                try:
                    # 一些Word版本会返回 Selection 而非 Range
                    test_range = getattr(rng, 'Range', None)
                    if test_range is not None:
                        rng = test_range
                except Exception:
                    pass

                inline_shapes = rng.InlineShapes
                inline_shape = inline_shapes.AddPicture(str(image_path))

                # 首选：将插入的 inline shape 转为浮于文字上的 Shape，并铺满整页
                try:
                    shp = inline_shape.ConvertToShape()

                    # 禁止锁定纵横比以便铺满整页
                    try:
                        shp.LockAspectRatio = False
                    except Exception:
                        pass

                    # 获取目标页的尺寸（优先使用 rng 的 section）
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
                        # 设置图片格式（浮于文字上方并铺满整页，参考前端/用户配置）
                        try:
                            # 使用数值常量以避免依赖命名常量差异
                            shp.WrapFormat.Type = 3
                            shp.RelativeHorizontalPosition = 1
                            shp.RelativeVerticalPosition = 1

                            # 尝试使用当前窗口的 Selection 对应的 section
                            try:
                                current_section = doc.ActiveWindow.Selection.Sections(1)
                            except Exception:
                                current_section = current_section

                            page_height = current_section.PageSetup.PageHeight
                            page_width = current_section.PageSetup.PageWidth

                            if page_height > page_width:
                                shp.Width = page_width
                                shp.Height = page_height
                            else:
                                shp.Width = page_height
                                shp.Height = page_width
                                try:
                                    shp.Rotation = 90
                                except Exception:
                                    try:
                                        shp.rotation = 90
                                    except Exception:
                                        pass

                            # 将Left/Top设置为较大的负值，确保在页面左上角覆盖
                            try:
                                shp.Left = -999995
                                shp.Top = -999995
                            except Exception:
                                pass

                            try:
                                shp.ZOrder(1)  # wdBringToFront = 1
                            except Exception:
                                pass
                        except Exception:
                            pass
                    else:
                        # 无法获取页设置时，回退为按比例缩放的 inline 插入
                        try:
                            if inline_shape.Width > 400:
                                ratio = 400 / inline_shape.Width
                                inline_shape.Width = 400
                                inline_shape.Height = int(inline_shape.Height * ratio)
                        except Exception:
                            pass

                except Exception as e:
                    # 如果 ConvertToShape 或属性设置失败，回退为 inline 插入并按可用宽度缩放
                    logger.warning(f"ConvertToShape 或浮动设置失败，回退到 inline 方式: {e}")
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
                            try:
                                if inline_shape.Width > usable_width:
                                    ratio = usable_width / inline_shape.Width
                                    inline_shape.Width = usable_width
                                    inline_shape.Height = int(inline_shape.Height * ratio)
                            except Exception:
                                pass
                    except Exception:
                        pass

                # 保存并关闭文档
                doc.SaveAs(str(output_path))
                doc.Close()
                logger.info(f"图片插入Word成功（Range方式，浮动/覆盖尝试）：{image_path} → {output_path}")
            except Exception as e:
                logger.error(f"通过 Range 插入图片失败：{e}", exc_info=True)
                try:
                    doc.Close(SaveChanges=False)
                except Exception:
                    pass
                raise
        except Exception as e:
            logger.error(f"插入图片失败：{word_path}", exc_info=True)
            raise WordProcessError(f"插入失败：{str(e)}") from e

