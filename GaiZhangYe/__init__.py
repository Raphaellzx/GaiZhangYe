"""
盖章页处理工具 - GaiZhangYe

一个专门用于批量处理Word文档和PDF文件中盖章页的Python应用程序。
提供Web界面操作方式，支持准备盖章页、盖章页覆盖和批量Word转PDF功能。
"""

__version__ = "0.1.0"
__author__ = "Your Name"
__email__ = "your@email.com"

# 导出核心模块
from GaiZhangYe.core import (
    stamp_prepare,
    stamp_overlay,
    batch_convert,
    data_communication
)

from GaiZhangYe.utils import (
    config,
    logger
)

from GaiZhangYe.web import (
    app
)

__all__ = [
    "stamp_prepare",
    "stamp_overlay",
    "batch_convert",
    "data_communication",
    "config",
    "logger",
    "app"
]
