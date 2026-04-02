import re
from pathlib import Path
from typing import List, Optional, Union
import os
import win32api
from GaiZhangYe.utils.logger import get_logger
from GaiZhangYe.core.models.exceptions import FileProcessError

logger = get_logger(__name__)

# 全局实例变量
_file_manager_instance = None


def get_file_manager() -> 'FileManager':
    """获取 FileManager 的单例实例"""
    global _file_manager_instance
    if _file_manager_instance is None:
        _file_manager_instance = FileManager()
    return _file_manager_instance


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


class FileManager:
    """
    文件管理类，提供文件和目录操作的统一接口
    支持功能1和功能2的路径管理，允许用户自定义路径
    """

    def __init__(self):
        self.logger = logger
        # 功能1目录配置
        self.func1_dirs = {
            "nostamped_word": None,
            "nostamped_pdf": None,
            "stamped_pages": None,
            "temp": None
        }
        # 功能2目录配置
        self.func2_dirs = {
            "images": None,
            "target_files": None,
            "result_word": None,
            "result_pdf": None
        }
        # 初始化默认路径
        self._init_default_dirs()

    def _init_default_dirs(self):
        """初始化默认目录路径"""
        # 获取用户文档目录
        if os.name == 'nt':  # Windows系统
            try:
                import ctypes
                CSIDL_PERSONAL = 5  # My Documents
                SHGFP_TYPE_CURRENT = 0  # Current value

                buf = ctypes.create_unicode_buffer(260)
                ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf)
                documents_path = Path(buf.value)
            except Exception as e:
                logger.error(f"获取文档目录失败: {e}")
                documents_path = Path.home() / "Documents"
        else:
            documents_path = Path.home() / "Documents"

        # 功能1默认路径
        func1_base_path = documents_path / "StampTool" / "Func1"
        self.func1_dirs["nostamped_word"] = func1_base_path / "Nostamped_Word"
        self.func1_dirs["nostamped_pdf"] = func1_base_path / "Nostamped_PDF"
        self.func1_dirs["stamped_pages"] = func1_base_path / "Stamped_Pages"
        self.func1_dirs["temp"] = func1_base_path / "Temp"

        # 功能2默认路径
        func2_base_path = documents_path / "StampTool" / "Func2"
        self.func2_dirs["images"] = func2_base_path / "Images"
        self.func2_dirs["target_files"] = func2_base_path / "Target_Files"
        self.func2_dirs["result_word"] = func2_base_path / "Result_Word"
        self.func2_dirs["result_pdf"] = func2_base_path / "Result_PDF"

        # 确保所有目录存在
        self._ensure_dirs_exist()

    def _ensure_dirs_exist(self):
        """确保所有目录存在"""
        for dir_path in self.func1_dirs.values():
            if dir_path:
                dir_path.mkdir(parents=True, exist_ok=True)
        for dir_path in self.func2_dirs.values():
            if dir_path:
                dir_path.mkdir(parents=True, exist_ok=True)

    # 功能1目录管理方法
    def set_custom_func1_dir(self, dir_type: str, dir_path: Path):
        """设置功能1的自定义目录路径"""
        if dir_type not in self.func1_dirs:
            raise FileProcessError(f"无效的目录类型: {dir_type}")
        if not isinstance(dir_path, Path):
            dir_path = Path(dir_path)
        # 确保目录存在
        dir_path.mkdir(parents=True, exist_ok=True)
        self.func1_dirs[dir_type] = dir_path
        logger.info(f"功能1目录 {dir_type} 已设置为: {dir_path}")

    def get_func1_dir(self, dir_type: str) -> Path:
        """获取功能1的目录路径"""
        if dir_type not in self.func1_dirs:
            raise FileProcessError(f"无效的目录类型: {dir_type}")
        dir_path = self.func1_dirs[dir_type]
        if not dir_path:
            raise FileProcessError(f"目录类型 {dir_type} 未设置")
        return dir_path

    # 功能2目录管理方法
    def set_custom_func2_dir(self, dir_type: str, dir_path: Path):
        """设置功能2的自定义目录路径"""
        if dir_type not in self.func2_dirs:
            raise FileProcessError(f"无效的目录类型: {dir_type}")
        if not isinstance(dir_path, Path):
            dir_path = Path(dir_path)
        # 确保目录存在
        dir_path.mkdir(parents=True, exist_ok=True)
        self.func2_dirs[dir_type] = dir_path
        logger.info(f"功能2目录 {dir_type} 已设置为: {dir_path}")

    def get_func2_dir(self, dir_type: str) -> Path:
        """获取功能2的目录路径"""
        if dir_type not in self.func2_dirs:
            raise FileProcessError(f"无效的目录类型: {dir_type}")
        dir_path = self.func2_dirs[dir_type]
        if not dir_path:
            raise FileProcessError(f"目录类型 {dir_type} 未设置")
        return dir_path

    def list_files(self, dir_path: Path, allowed_extensions: Optional[List[str]] = None) -> List[Path]:
        """
        列出目录下的所有文件（支持按扩展名过滤）
        :param dir_path: 目录路径
        :param allowed_extensions: 允许的扩展名列表，如[".docx", ".doc"]
        :return: 符合条件的文件路径列表
        """
        if not dir_path.exists() or not dir_path.is_dir():
            raise FileProcessError(f"目录不存在或不是目录：{dir_path}")

        all_files = []
        for item in dir_path.iterdir():
            if item.is_file():
                if allowed_extensions:
                    if item.suffix.lower() in [ext.lower() for ext in allowed_extensions]:
                        all_files.append(item)
                else:
                    all_files.append(item)

        return sorted(all_files)

    def check_file_exists(self, file_path: Path) -> bool:
        """
        检查文件是否存在
        :param file_path: 文件路径
        :return: 文件是否存在的布尔值
        """
        return file_path.exists() and file_path.is_file()

    def get_file_size(self, file_path: Path) -> int:
        """
        获取文件大小（字节）
        :param file_path: 文件路径
        :return: 文件大小（字节）
        """
        if not self.check_file_exists(file_path):
            raise FileProcessError(f"文件不存在：{file_path}")
        return file_path.stat().st_size

    def create_directory(self, dir_path: Path) -> None:
        """
        创建目录（如果不存在）
        :param dir_path: 目录路径
        """
        dir_path.mkdir(parents=True, exist_ok=True)

    def delete_file(self, file_path: Path) -> None:
        """
        删除文件
        :param file_path: 文件路径
        """
        if self.check_file_exists(file_path):
            file_path.unlink()
            logger.debug(f"文件已删除：{file_path}")

    def delete_directory(self, dir_path: Path) -> None:
        """
        删除目录（递归删除所有内容）
        :param dir_path: 目录路径
        """
        if dir_path.exists() and dir_path.is_dir():
            import shutil
            shutil.rmtree(dir_path)
            logger.debug(f"目录已删除：{dir_path}")

    def get_available_disk_space(self, dir_path: Path) -> int:
        """
        获取目录所在驱动器的可用磁盘空间（字节）
        :param dir_path: 目录路径
        :return: 可用磁盘空间（字节）
        """
        try:
            free_bytes = win32api.GetDiskFreeSpaceEx(str(dir_path))[0]
            return free_bytes
        except Exception as e:
            logger.error(f"获取磁盘空间失败：{e}")
            raise FileProcessError(f"无法获取磁盘空间：{e}")

    def is_file_accessible(self, file_path: Path) -> bool:
        """
        检查文件是否可访问（读/写权限）
        :param file_path: 文件路径
        :return: 文件是否可访问的布尔值
        """
        try:
            # 检查读权限
            if not os.access(file_path, os.R_OK):
                return False

            # 检查写权限
            if not os.access(file_path, os.W_OK):
                return False

            return True
        except Exception:
            return False

    def get_file_creation_time(self, file_path: Path) -> int:
        """
        获取文件创建时间（时间戳）
        :param file_path: 文件路径
        :return: 文件创建时间（时间戳）
        """
        if not self.check_file_exists(file_path):
            raise FileProcessError(f"文件不存在：{file_path}")
        return file_path.stat().st_ctime

    def get_file_modification_time(self, file_path: Path) -> int:
        """
        获取文件修改时间（时间戳）
        :param file_path: 文件路径
        :return: 文件修改时间（时间戳）
        """
        if not self.check_file_exists(file_path):
            raise FileProcessError(f"文件不存在：{file_path}")
        return file_path.stat().st_mtime

    def get_file_access_time(self, file_path: Path) -> int:
        """
        获取文件访问时间（时间戳）
        :param file_path: 文件路径
        :return: 文件访问时间（时间戳）
        """
        if not self.check_file_exists(file_path):
            raise FileProcessError(f"文件不存在：{file_path}")
        return file_path.stat().st_atime

    def move_file(self, source_path: Path, destination_path: Path) -> None:
        """
        移动文件
        :param source_path: 源文件路径
        :param destination_path: 目标文件路径
        """
        if not self.check_file_exists(source_path):
            raise FileProcessError(f"源文件不存在：{source_path}")

        import shutil
        shutil.move(source_path, destination_path)
        logger.debug(f"文件已移动：{source_path} -> {destination_path}")

    def copy_file(self, source_path: Path, destination_path: Path) -> None:
        """
        复制文件
        :param source_path: 源文件路径
        :param destination_path: 目标文件路径
        """
        if not self.check_file_exists(source_path):
            raise FileProcessError(f"源文件不存在：{source_path}")

        import shutil
        shutil.copy(source_path, destination_path)
        logger.debug(f"文件已复制：{source_path} -> {destination_path}")
