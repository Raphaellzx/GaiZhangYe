import logging
import logging.handlers
import os
from pathlib import Path
from typing import Optional

# 导入项目配置（确保config.py已实现）
from GaiZhangYe.utils.config import get_settings

# 全局日志格式（包含时间、模块、行号，便于定位问题）
DEFAULT_LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s"
DEFAULT_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"

# 日志级别映射（将.env中的字符串级别转为logging常量）
LOG_LEVEL_MAP = {
    "DEBUG": logging.DEBUG,
    "INFO": logging.INFO,
    "WARNING": logging.WARNING,
    "ERROR": logging.ERROR,
    "CRITICAL": logging.CRITICAL,
}

def _get_log_level() -> int:
    """从配置中获取日志级别（默认INFO）"""
    settings = get_settings()
    level_str = settings.log_level.strip().upper()
    return LOG_LEVEL_MAP.get(level_str, logging.INFO)

def _ensure_log_dir() -> Path:
    """确保日志目录存在，返回日志目录路径"""
    settings = get_settings()
    log_dir = Path(settings.log_dir)
    log_dir.mkdir(exist_ok=True, parents=True)
    return log_dir

def get_logger(name: Optional[str] = None) -> logging.Logger:
    """
    获取全局统一配置的logger实例
    :param name: logger名称（建议传__name__，即模块路径）
    :return: 配置好的logger
    """
    # 避免重复添加handler（关键：防止多模块调用时重复输出日志）
    logger_name = name or "GaiZhangYe"
    logger = logging.getLogger(logger_name)
    
    if logger.handlers:
        return logger

    # 1. 设置基础日志级别
    logger.setLevel(_get_log_level())
    logger.propagate = False  # 禁止向上传播（避免root logger重复输出）

    # 2. 创建格式器
    formatter = logging.Formatter(DEFAULT_LOG_FORMAT, datefmt=DEFAULT_DATE_FORMAT)

    # 3. 控制台Handler（开发环境实时查看）
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    console_handler.setLevel(_get_log_level())
    logger.addHandler(console_handler)

    # 4. 主日志文件Handler（按大小轮转，避免文件过大）
    log_dir = _ensure_log_dir()
    main_log_file = log_dir / "app.log"
    file_handler = logging.handlers.RotatingFileHandler(
        main_log_file,
        maxBytes=10 * 1024 * 1024,  # 单个文件最大10MB
        backupCount=5,              # 保留5个备份文件
        encoding="utf-8",           # 确保中文正常显示
    )
    file_handler.setFormatter(formatter)
    file_handler.setLevel(_get_log_level())
    logger.addHandler(file_handler)

    # 5. 错误日志单独分离（仅记录ERROR/CRITICAL级别，便于排查问题）
    error_log_file = log_dir / "error.log"
    error_handler = logging.handlers.RotatingFileHandler(
        error_log_file,
        maxBytes=5 * 1024 * 1024,   # 单个错误日志最大5MB
        backupCount=3,              # 保留3个备份
        encoding="utf-8",
    )
    error_handler.setFormatter(formatter)
    error_handler.setLevel(logging.ERROR)  # 仅记录错误及以上级别
    logger.addHandler(error_handler)

    return logger

# 便捷导出：项目全局通用logger（非必需，推荐各模块用__name__）
global_logger = get_logger("GaiZhangYe")