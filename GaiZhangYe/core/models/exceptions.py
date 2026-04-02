# GaiZhangYe/core/models/exceptions.py
"""
业务异常定义：统一管理项目中所有自定义异常，提供详细的错误信息和分类
"""

class BusinessError(Exception):
    """业务逻辑异常基类，所有业务相关异常都应该继承自该类"""
    def __init__(self, message: str, details: dict = None):
        """
        初始化业务异常
        :param message: 错误信息
        :param details: 错误详情（可选，如文件名、路径、页码等）
        """
        super().__init__(message)
        self.details = details or {}
        self.timestamp = __import__('datetime').datetime.now()

    def __str__(self):
        """返回包含详细信息的字符串表示"""
        base_msg = super().__str__()
        if self.details:
            details_str = ", ".join(f"{k}={v}" for k, v in self.details.items())
            return f"{base_msg} ({details_str})"
        return base_msg


class WordProcessError(BusinessError):
    """Word处理相关异常，包括Word文件读取、写入、转换等错误"""
    def __init__(self, message: str, file_path: str = None, page_number: int = None):
        details = {}
        if file_path:
            details["file_path"] = file_path
        if page_number is not None:
            details["page_number"] = page_number
        super().__init__(message, details)


class PdfProcessError(BusinessError):
    """PDF处理相关异常，包括PDF文件读取、写入、转换、提取等错误"""
    def __init__(self, message: str, file_path: str = None, page_number: int = None):
        details = {}
        if file_path:
            details["file_path"] = file_path
        if page_number is not None:
            details["page_number"] = page_number
        super().__init__(message, details)


class ImageProcessError(BusinessError):
    """图片处理相关异常，包括图片读取、写入、缩放、转换等错误"""
    def __init__(self, message: str, file_path: str = None, width: int = None, height: int = None):
        details = {}
        if file_path:
            details["file_path"] = file_path
        if width is not None:
            details["width"] = width
        if height is not None:
            details["height"] = height
        super().__init__(message, details)


class FileProcessError(BusinessError):
    """文件处理相关异常，包括文件读取、写入、删除、复制等错误"""
    def __init__(self, message: str, file_path: str = None, operation: str = None):
        details = {}
        if file_path:
            details["file_path"] = file_path
        if operation:
            details["operation"] = operation
        super().__init__(message, details)


class DirCreateError(BusinessError):
    """目录创建异常，包括目录创建失败、权限不足等错误"""
    def __init__(self, message: str, dir_path: str = None, reason: str = None):
        details = {}
        if dir_path:
            details["dir_path"] = dir_path
        if reason:
            details["reason"] = reason
        super().__init__(message, details)
