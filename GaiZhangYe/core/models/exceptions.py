# GaiZhangYe/core/models/exceptions.py
"""
业务异常定义：统一管理项目中所有自定义异常
"""

class BusinessError(Exception):
    """业务逻辑异常基类"""
    def __init__(self, message: str):
        super().__init__(message)

class WordProcessError(BusinessError):
    """Word处理相关异常"""
    pass

class PdfProcessError(BusinessError):
    """PDF处理相关异常"""
    pass

class ImageProcessError(BusinessError):
    """图片处理相关异常"""
    pass

class FileProcessError(BusinessError):
    """文件处理相关异常"""
    pass

class DirCreateError(BusinessError):
    """目录创建异常"""
    pass
