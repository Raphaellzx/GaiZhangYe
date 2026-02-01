#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
core层数据沟通模块
实现前后端通过文件进行数据交换的功能
"""

import json
from typing import Dict, Any
from GaiZhangYe.core.basic.file_manager import get_file_manager
from GaiZhangYe.core.basic.file_processor import FileProcessor
from GaiZhangYe.utils.logger import get_logger

logger = get_logger(__name__)


class DataCommunicationService:
    """数据沟通服务类"""

    def __init__(self):
        # 创建文件管理器和处理器实例（使用单例）
        self.file_manager = get_file_manager()
        self.file_processor = FileProcessor()

        # 定义数据文件路径（使用文件管理器获取正确的路径）
        self.func1_data_file = self.file_manager.get_func1_dir("temp") / 'target_pages.json'
        self.func2_data_file = self.file_manager.get_func2_dir("temp") / 'stamp_config.json'

    def get_func1_data(self) -> Dict[str, Any]:
        """获取func1的target_pages数据"""
        try:
            if self.func1_data_file.exists():
                with open(self.func1_data_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return {}
        except Exception as e:
            logger.error(f"获取func1数据失败: {str(e)}", exc_info=True)
            return {}

    def save_func1_data(self, data: Dict[str, Any]) -> bool:
        """保存func1的target_pages数据"""
        try:
            # 确保目录存在
            self.func1_data_file.parent.mkdir(exist_ok=True)
            with open(self.func1_data_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            logger.info(f"func1数据已保存到: {self.func1_data_file}")
            return True
        except Exception as e:
            logger.error(f"保存func1数据失败: {str(e)}", exc_info=True)
            return False

    def get_func2_data(self) -> Dict[str, Any]:
        """获取func2的stamp_config数据"""
        try:
            if self.func2_data_file.exists():
                with open(self.func2_data_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return {}
        except Exception as e:
            logger.error(f"获取func2数据失败: {str(e)}", exc_info=True)
            return {}

    def save_func2_data(self, data: Dict[str, Any]) -> bool:
        """保存func2的stamp_config数据"""
        try:
            # 确保目录存在
            self.func2_data_file.parent.mkdir(exist_ok=True)
            with open(self.func2_data_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
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

