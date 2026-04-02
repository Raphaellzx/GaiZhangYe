#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
core层数据沟通模块
实现前后端通过文件进行数据交换的功能，使用类型安全的数据模型
"""

import json
from typing import Dict, Any, Optional
from pathlib import Path
from GaiZhangYe.core.basic.file_manager import get_file_manager
from GaiZhangYe.utils.logger import get_logger
from GaiZhangYe.core.models.data_models import Func1Data, Func2Data, StampConfig

logger = get_logger(__name__)


class DataCommunicationService:
    """数据沟通服务类：使用类型安全的数据模型"""

    def __init__(self):
        # 创建文件管理器实例（使用单例）
        self.file_manager = get_file_manager()

        # 初始化数据目录
        self._init_directories()

    def _init_directories(self):
        """初始化数据目录"""
        # 使用文件管理器获取功能1和功能2的默认数据目录
        self.func1_data_dir = self.file_manager.get_func1_dir("temp") / "data"
        self.func2_data_dir = self.file_manager.get_func2_dir("images").parent / "data"

        # 数据文件路径
        self.func1_data_file = self.func1_data_dir / 'target_pages.json'
        self.func2_data_file = self.func2_data_dir / 'stamp_config.json'

        # 确保目录存在
        self.func1_data_dir.mkdir(parents=True, exist_ok=True)
        self.func2_data_dir.mkdir(parents=True, exist_ok=True)

    def get_func1_data(self) -> Func1Data:
        """获取func1的target_pages数据（类型安全）"""
        try:
            if self.func1_data_file.exists():
                with open(self.func1_data_file, 'r', encoding='utf-8') as f:
                    raw_data = json.load(f)
                return Func1Data(**raw_data)
            return Func1Data(target_pages={})
        except Exception as e:
            logger.error(f"获取func1数据失败: {str(e)}", exc_info=True)
            return Func1Data(target_pages={})

    def save_func1_data(self, data: Func1Data) -> bool:
        """保存func1的target_pages数据（类型安全）"""
        try:
            # 确保data是Func1Data类型
            if not isinstance(data, Func1Data):
                data = Func1Data(**data)

            with open(self.func1_data_file, 'w', encoding='utf-8') as f:
                json.dump(data.__dict__, f, ensure_ascii=False, indent=2)
            logger.info(f"func1数据已保存到: {self.func1_data_file}")
            return True
        except Exception as e:
            logger.error(f"保存func1数据失败: {str(e)}", exc_info=True)
            return False

    def get_func2_data(self) -> Func2Data:
        """获取func2的stamp_config数据（类型安全）"""
        try:
            if self.func2_data_file.exists():
                with open(self.func2_data_file, 'r', encoding='utf-8') as f:
                    raw_data = json.load(f)
                # 将字典转换为StampConfig对象
                configs = []
                if "configs" in raw_data:
                    for config in raw_data["configs"]:
                        try:
                            configs.append(StampConfig(**config))
                        except Exception as e:
                            logger.error(f"解析配置失败: {str(e)}", exc_info=True)
                return Func2Data(configs=configs)
            return Func2Data(configs=[])
        except Exception as e:
            logger.error(f"获取func2数据失败: {str(e)}", exc_info=True)
            return Func2Data(configs=[])

    def save_func2_data(self, data: Func2Data) -> bool:
        """保存func2的stamp_config数据（类型安全）"""
        try:
            # 确保data是Func2Data类型
            if not isinstance(data, Func2Data):
                data = Func2Data(**data)

            # 将StampConfig对象转换为字典
            configs_list = []
            for config in data.configs:
                configs_list.append(config.__dict__)
            with open(self.func2_data_file, 'w', encoding='utf-8') as f:
                json.dump({"configs": configs_list}, f, ensure_ascii=False, indent=2)
            logger.info(f"func2数据已保存到: {self.func2_data_file}")
            return True
        except Exception as e:
            logger.error(f"保存func2数据失败: {str(e)}", exc_info=True)
            return False


# 单例模式
_data_service = None


def get_data_service() -> DataCommunicationService:
    """获取数据服务实例"""
    global _data_service
    if _data_service is None:
        _data_service = DataCommunicationService()
    return _data_service

