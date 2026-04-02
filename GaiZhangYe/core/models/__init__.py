"""数据模型模块：定义项目中所有业务数据结构和配置模型"""

from .data_models import (
    Func1Data, PageConfig, Func2Data, StampConfig, ImageLocationType,
    ProcessResult, WordFileInfo, PdfFileInfo, ImageFileInfo
)
from .exceptions import (
    BusinessError, WordProcessError, PdfProcessError, ImageProcessError,
    FileProcessError, DirCreateError
)

__all__ = [
    "Func1Data", "PageConfig", "Func2Data", "StampConfig", "ImageLocationType",
    "ProcessResult", "WordFileInfo", "PdfFileInfo", "ImageFileInfo",
    "BusinessError", "WordProcessError", "PdfProcessError", "ImageProcessError",
    "FileProcessError", "DirCreateError"
]
