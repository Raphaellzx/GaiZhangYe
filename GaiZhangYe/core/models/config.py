# GaiZhangYe/core/models/config.py
"""
业务目录配置模型
"""
from dataclasses import dataclass


@dataclass
class BusinessDirConfig:
    """业务目录配置"""
    func1_nostamped_word: str
    func1_nostamped_pdf: str
    func1_stamped_pages: str

    func2_images: str
    func2_target_files: str
    func2_result_word: str
    func2_result_pdf: str


# 可以在这里添加其他业务配置模型
