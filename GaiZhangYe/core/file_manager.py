# GaiZhangYe/core/file_manager.py
from pathlib import Path
from typing import Dict, Optional
from GaiZhangYe.utils.logger import get_logger
from GaiZhangYe.core.models.exceptions import DirCreateError

logger = get_logger(__name__)

class FileManager:
    """业务目录管理器：创建/管理business_data下的所有目录"""
    def __init__(self, root_dir: Optional[Path] = None):
        # 业务根目录默认：项目根/business_data
        self.root_dir = root_dir or Path(__file__).parent.parent.parent.parent / "business_data"
        self.func1_dirs: Dict[str, Path] = {}  # 功能1目录映射
        self.func2_dirs: Dict[str, Path] = {}  # 功能2目录映射
        self._init_all_dirs()

    def _init_all_dirs(self):
        """初始化所有业务目录（func1/func2）"""
        try:
            # 功能1目录：Nostamped_Word/Nostamped_PDF/Stamped_Pages
            func1_root = self.root_dir / "func1"
            self.func1_dirs = {
                "nostamped_word": func1_root / "Nostamped_Word",
                "nostamped_pdf": func1_root / "Nostamped_PDF",
                "stamped_pages": func1_root / "Stamped_Pages"
            }
            # 功能2目录：Images/TargetFiles/Result_Word/Result_PDF
            func2_root = self.root_dir / "func2"
            self.func2_dirs = {
                "images": func2_root / "Images",
                "target_files": func2_root / "TargetFiles",
                "result_word": func2_root / "Result_Word",
                "result_pdf": func2_root / "Result_PDF"
            }
            # 批量创建所有目录
            all_dirs = list(self.func1_dirs.values()) + list(self.func2_dirs.values())
            for dir_path in all_dirs:
                dir_path.mkdir(exist_ok=True, parents=True)
                logger.info(f"业务目录初始化完成：{dir_path}")
        except Exception as e:
            logger.error("业务目录初始化失败", exc_info=True)
            raise DirCreateError(f"创建业务目录失败：{str(e)}") from e

    def get_func1_dir(self, dir_type: str) -> Path:
        """获取功能1指定目录（dir_type: nostamped_word/nostamped_pdf/stamped_pages）"""
        if dir_type not in self.func1_dirs:
            raise ValueError(f"功能1无此目录类型：{dir_type}")
        return self.func1_dirs[dir_type]

    def get_func2_dir(self, dir_type: str) -> Path:
        """获取功能2指定目录（dir_type: images/target_files/result_word/result_pdf）"""
        if dir_type not in self.func2_dirs:
            raise ValueError(f"功能2无此目录类型：{dir_type}")
        return self.func2_dirs[dir_type]

    def clean_dir(self, dir_path: Path, keep_latest: int = 0):
        """清理目录（保留最新N个文件，默认全清）"""
        if not dir_path.exists():
            return
        # 按修改时间排序文件
        files = sorted(dir_path.glob("*"), key=lambda f: f.stat().st_mtime, reverse=True)
        # 保留最新N个，删除其余
        for file in files[keep_latest:]:
            if file.is_file():
                file.unlink()
                logger.debug(f"清理过期文件：{file}")