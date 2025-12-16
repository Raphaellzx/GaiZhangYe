# GaiZhangYe/utils/config.py
from pathlib import Path
from pydantic_settings import BaseSettings, SettingsConfigDict
from GaiZhangYe.utils.logger import get_logger

logger = get_logger(__name__)

class AppSettings(BaseSettings):
    """全局配置（从.env加载）"""
    # 日志配置
    log_level: str = "INFO"
    log_dir: Path = Path(__file__).parent.parent.parent / "logs"
    
    # 业务目录配置（可通过.env自定义，默认用项目根/business_data）
    business_data_root: Path = Path(__file__).parent.parent.parent / "business_data"

    # 加载.env文件
    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        case_sensitive=False
    )

# 单例配置
_settings = AppSettings()

def get_settings() -> AppSettings:
    return _settings

# 初始化业务目录（启动时自动执行）
def init_business_dirs():
    from GaiZhangYe.core.file_manager import FileManager
    try:
        FileManager(root_dir=_settings.business_data_root)
        logger.info(f"业务目录初始化完成：{_settings.business_data_root}")
    except Exception as e:
        logger.error("业务目录初始化失败", exc_info=True)
        raise