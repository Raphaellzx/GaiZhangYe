from dataclasses import dataclass, field
from enum import Enum
from typing import List, Dict, Optional, Any, Union
from pathlib import Path
from datetime import datetime


class ImageLocationType(Enum):
    """图片插入位置类型"""
    PAGE_NUMBER = 1
    LAST_PAGE = 2


@dataclass
class PageConfig:
    """页面配置数据模型"""
    filename: str
    pages: List[int] = field(metadata={"description": "至少需要指定一个页面"})

    def __post_init__(self):
        if not self.pages or len(self.pages) == 0:
            raise ValueError("页面配置至少需要指定一个页面")
        for page_num in self.pages:
            if not isinstance(page_num, int) or page_num < 1:
                raise ValueError(f"无效的页码: {page_num}，必须是大于0的整数")


@dataclass
class Func1Data:
    """功能1（准备盖章页）数据模型"""
    target_pages: Dict[str, List[int]] = field(default_factory=dict)

    def __post_init__(self):
        if not isinstance(self.target_pages, dict):
            raise ValueError("target_pages必须是字典类型")
        for filename, pages in self.target_pages.items():
            if not isinstance(filename, str) or not filename:
                raise ValueError("文件名必须是有效的字符串")
            if not isinstance(pages, list) or len(pages) == 0:
                raise ValueError(f"{filename}的页面配置不能为空")
            for page_num in pages:
                if not isinstance(page_num, int) or page_num < 1:
                    raise ValueError(f"{filename}中存在无效的页码: {page_num}")


@dataclass
class StampConfig:
    """盖章配置数据模型"""
    filename: str
    image_files: List[str] = field(default_factory=list)
    positions: List[Union[int, str]] = field(default_factory=list)
    image_width: int = 100

    def __post_init__(self):
        if not isinstance(self.filename, str) or not self.filename.strip():
            raise ValueError("文件名必须是有效的字符串")
        if not self.image_files:
            raise ValueError("盖章配置至少需要指定一个图片文件")
        if not self.positions:
            raise ValueError("盖章配置至少需要指定一个插入位置")
        for img_path in self.image_files:
            if not isinstance(img_path, str) or not img_path:
                raise ValueError("图片路径必须是有效的字符串")
        for pos in self.positions:
            if not (isinstance(pos, int) and pos > 0 or isinstance(pos, str) and pos.strip()):
                raise ValueError(f"无效的插入位置: {pos}")
        if not isinstance(self.image_width, int) or self.image_width <= 0:
            raise ValueError(f"无效的图片宽度: {self.image_width}，必须是正整数")


@dataclass
class Func2Data:
    """功能2（盖章页覆盖）数据模型"""
    configs: List[StampConfig] = field(default_factory=list)

    def __post_init__(self):
        if not isinstance(self.configs, list):
            raise ValueError("configs必须是列表类型")
        for config in self.configs:
            if not isinstance(config, StampConfig):
                raise ValueError("configs列表中的元素必须是StampConfig类型")


@dataclass
class WordFileInfo:
    """Word文件信息数据模型"""
    file_path: Path
    page_count: int = 0
    file_size: int = 0
    last_modified: Optional[datetime] = None

    def __post_init__(self):
        if not isinstance(self.file_path, Path):
            raise ValueError("file_path必须是Path类型")
        if not self.file_path.exists():
            raise FileNotFoundError(f"文件不存在: {self.file_path}")
        if self.page_count < 0:
            raise ValueError("页码数不能小于0")
        if self.file_size < 0:
            raise ValueError("文件大小不能小于0")


@dataclass
class PdfFileInfo:
    """PDF文件信息数据模型"""
    file_path: Path
    page_count: int = 0
    file_size: int = 0
    last_modified: Optional[datetime] = None

    def __post_init__(self):
        if not isinstance(self.file_path, Path):
            raise ValueError("file_path必须是Path类型")
        if not self.file_path.exists():
            raise FileNotFoundError(f"文件不存在: {self.file_path}")
        if self.page_count < 0:
            raise ValueError("页码数不能小于0")
        if self.file_size < 0:
            raise ValueError("文件大小不能小于0")


@dataclass
class ImageFileInfo:
    """图片文件信息数据模型"""
    file_path: Path
    width: int = 0
    height: int = 0
    file_size: int = 0
    last_modified: Optional[datetime] = None

    def __post_init__(self):
        if not isinstance(self.file_path, Path):
            raise ValueError("file_path必须是Path类型")
        if not self.file_path.exists():
            raise FileNotFoundError(f"文件不存在: {self.file_path}")
        if self.width < 0 or self.height < 0:
            raise ValueError("图片尺寸不能小于0")
        if self.file_size < 0:
            raise ValueError("文件大小不能小于0")


@dataclass
class ProcessResult:
    """处理结果数据模型"""
    success: bool = False
    processed_files: List[Path] = field(default_factory=list)
    failed_files: List[Dict[str, Any]] = field(default_factory=list)
    total_time: float = 0.0
    message: str = ""

    def __post_init__(self):
        if not isinstance(self.processed_files, list):
            raise ValueError("processed_files必须是列表类型")
        for file_path in self.processed_files:
            if not isinstance(file_path, Path):
                raise ValueError("processed_files列表中的元素必须是Path类型")
        if not isinstance(self.failed_files, list):
            raise ValueError("failed_files必须是列表类型")
        for failed in self.failed_files:
            if not isinstance(failed, dict) or "file" not in failed or "error" not in failed:
                raise ValueError("failed_files列表中的元素必须包含'file'和'error'键")
        if self.total_time < 0:
            raise ValueError("处理时间不能小于0")


@dataclass
class ProcessingContext:
    """处理上下文数据模型"""
    task_id: str
    start_time: datetime
    end_time: Optional[datetime] = None
    parameters: Dict[str, Any] = field(default_factory=dict)
    environment: Dict[str, Any] = field(default_factory=dict)

    def __post_init__(self):
        if not isinstance(self.task_id, str) or not self.task_id:
            raise ValueError("task_id必须是有效的字符串")
        if not isinstance(self.start_time, datetime):
            raise ValueError("start_time必须是datetime类型")
        if self.end_time and self.end_time < self.start_time:
            raise ValueError("end_time不能早于start_time")
