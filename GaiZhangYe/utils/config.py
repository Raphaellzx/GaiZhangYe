# utils/config.py
import logging
from pathlib import Path
from pydantic_settings import BaseSettings, SettingsConfigDict

class AppSettings(BaseSettings):
    """全局配置模型（从.env加载）"""
    # 日志配置（与.env.example对应）
    log_level: str = "INFO"
    log_dir: Path = Path(__file__).parent.parent.parent / "logs"  # 默认项目根/logs
    
    # 业务目录配置
    business_data_root: Path = Path(__file__).parent.parent.parent / "business_data"
    
    # 图片默认缩放宽度（功能2）
    image_default_width: int = 800

    # 加载.env文件
    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        case_sensitive=False,  # 不区分大小写（如LOG_LEVEL和log_level都可）
        extra="ignore"         # 忽略.env中未定义的配置项
    )

# 单例配置（全局复用）
_settings = AppSettings()

def get_settings() -> AppSettings:
    return _settings

# 初始化业务目录（启动时自动执行）
def init_business_dirs():
    from GaiZhangYe.core.file_manager import FileManager
    try:
        FileManager(root_dir=_settings.business_data_root)
    except Exception as e:
        # 初始化失败时用临时logger输出（避免循环依赖）
        temp_logger = logging.getLogger("GaiZhangYe-init")
        temp_logger.error(f"业务目录初始化失败：{e}", exc_info=True)
        raise