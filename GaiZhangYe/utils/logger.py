# your_project/utils/logger.py
import logging
import logging.handlers
import os
from pathlib import Path
from typing import Optional

from GaiZhangYe.utils.config import get_settings  # 加载配置（如日志级别、路径）

# 确保日志目录存在
LOG_DIR = Path(get_settings().log_dir)
LOG_DIR.mkdir(exist_ok=True, parents=True)

# 日志格式（包含时间、模块、级别、消息，便于排查）
LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s"
DATE_FORMAT = "%Y-%m-%d %H:%M:%S"

def get_logger(name: Optional[str] = None) -> logging.Logger:
    """获取统一配置的logger（各模块调用此函数即可）"""
    # 避免重复添加handler
    logger = logging.getLogger(name or __name__)
    if logger.handlers:
        return logger

    # 基础配置
    logger.setLevel(get_settings().log_level)
    formatter = logging.Formatter(LOG_FORMAT, datefmt=DATE_FORMAT)

    # 1. 控制台handler（开发环境）
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # 2. 文件handler（按大小轮转，避免日志文件过大）
    file_handler = logging.handlers.RotatingFileHandler(
        LOG_DIR / "app.log",
        maxBytes=10 * 1024 * 1024,  # 10MB
        backupCount=5,  # 保留5个备份
        encoding="utf-8"
    )
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    # 3. 错误日志单独分离（可选，便于排查）
    error_handler = logging.handlers.RotatingFileHandler(
        LOG_DIR / "error.log",
        maxBytes=5 * 1024 * 1024,
        backupCount=3,
        encoding="utf-8"
    )
    error_handler.setLevel(logging.ERROR)
    error_handler.setFormatter(formatter)
    logger.addHandler(error_handler)

    return logger