# GaiZhangYe/core/file_processor.py
"""
通用文件处理工具：提供文件校验、遍历、清理等通用功能
"""
from pathlib import Path
from typing import List
from GaiZhangYe.utils.logger import get_logger
from GaiZhangYe.core.models.exceptions import FileProcessError
import re
from typing import Union

logger = get_logger(__name__)


def windows_natural_sort_key(filename: Union[str, Path]) -> List:
    """
    将文件名转换为 Windows 资源管理器风格的排序键
    可用于 sorted(..., key=windows_natural_sort_key)
    """
    if isinstance(filename, Path):
        filename = filename.name

    parts = []
    for text in re.split(r'(\d+)', str(filename)):
        if text.isdigit():
            parts.append(int(text))
        else:
            parts.append(text.lower())
    return parts


def sort_files_windows_style(files: List[Union[str, Path]]) -> List[Union[str, Path]]:
    return sorted(files, key=windows_natural_sort_key)


def sort_dicts_by_name_windows_style(dicts: List[dict], name_key: str = "name") -> List[dict]:
    return sorted(dicts, key=lambda d: windows_natural_sort_key(d.get(name_key, "")))


class FileProcessor:
    """通用文件处理器"""

    def list_files(self, dir_path: Path, allowed_extensions: List[str] = None) -> List[Path]:
        """
        列出目录下的所有文件（支持按扩展名过滤）
        :param dir_path: 目录路径
        :param allowed_extensions: 允许的扩展名列表，如[".docx", ".doc"]
        :return: 符合条件的文件路径列表
        """
        if not dir_path.exists() or not dir_path.is_dir():
            raise FileProcessError(f"目录不存在或不是目录：{dir_path}")

        # 获取所有文件
        all_files = [f for f in dir_path.iterdir() if f.is_file()]

        # 按扩展名过滤
        if allowed_extensions:
            allowed_extensions = [ext.lower() for ext in allowed_extensions]
            filtered_files = [
                f for f in all_files
                if f.suffix.lower() in allowed_extensions
            ]
        else:
            filtered_files = all_files

        logger.debug(f"在目录{dir_path}中找到{len(filtered_files)}个文件")
        return filtered_files

    def check_file_exists(self, file_path: Path) -> bool:
        """
        检查文件是否存在
        :param file_path: 文件路径
        :return: 存在返回True，否则返回False
        """
        return file_path.exists() and file_path.is_file()

    def check_file_type(self, file_path: Path, allowed_extensions: List[str]) -> bool:
        """
        检查文件类型是否符合要求
        :param file_path: 文件路径
        :param allowed_extensions: 允许的扩展名列表
        :return: 符合返回True，否则返回False
        """
        if not self.check_file_exists(file_path):
            return False

        allowed_extensions = [ext.lower() for ext in allowed_extensions]
        return file_path.suffix.lower() in allowed_extensions

    def get_file_size(self, file_path: Path) -> int:
        """
        获取文件大小（字节）
        :param file_path: 文件路径
        :return: 文件大小（字节）
        """
        if not self.check_file_exists(file_path):
            raise FileProcessError(f"文件不存在：{file_path}")

        return file_path.stat().st_size

    def delete_file(self, file_path: Path) -> None:
        """
        删除文件
        :param file_path: 文件路径
        """
        if self.check_file_exists(file_path):
            file_path.unlink()
            logger.debug(f"文件已删除：{file_path}")

    def batch_delete_files(self, file_paths: List[Path]) -> None:
        """
        批量删除文件
        :param file_paths: 文件路径列表
        """
        for file_path in file_paths:
            self.delete_file(file_path)
