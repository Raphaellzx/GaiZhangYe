#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
core层数据沟通模块
实现前后端通过文件进行数据交换的功能
"""

import json
from pathlib import Path
from typing import Dict, Any, Optional
from GaiZhangYe.core.basic.file_manager import get_file_manager
from GaiZhangYe.core.basic.file_processor import FileProcessor
from GaiZhangYe.core.basic.file_processor import windows_natural_sort_key


class DataCommunicationService:
    """数据沟通服务类"""
    
    def __init__(self):
        # 定义数据文件路径
        self.func1_data_file = Path(__file__).parent.parent / 'business_data' / 'func1' / '.temp' / 'target_pages.json'
        self.func2_data_file = Path(__file__).parent.parent / 'business_data' / 'func2' / '.temp' / 'stamp_config.json'
        
        # 创建文件管理器和处理器实例（使用单例）
        self.file_manager = get_file_manager()
        self.file_processor = FileProcessor()
    
    def get_func1_data(self) -> Dict[str, Any]:
        """获取func1的target_pages数据"""
        try:
            if self.func1_data_file.exists():
                with open(self.func1_data_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return {}
        except Exception as e:
            # 处理异常
            return {}
    
    def save_func1_data(self, data: Dict[str, Any]) -> bool:
        """保存func1的target_pages数据"""
        try:
            # 确保目录存在
            self.func1_data_file.parent.mkdir(exist_ok=True)
            with open(self.func1_data_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            return False
    
    def get_func2_data(self) -> Dict[str, Any]:
        """获取func2的stamp_config数据"""
        try:
            if self.func2_data_file.exists():
                with open(self.func2_data_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return {}
        except Exception as e:
            return {}
    
    def save_func2_data(self, data: Dict[str, Any]) -> bool:
        """保存func2的stamp_config数据"""
        try:
            # 确保目录存在
            self.func2_data_file.parent.mkdir(exist_ok=True)
            with open(self.func2_data_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            return False
    
    def scan_business_data(self) -> bool:
        """扫描business_data目录并生成默认数据（调用具体子方法）"""
        try:
            ok1 = self.scan_func1()
            ok2 = self.scan_func2()
            return bool(ok1 and ok2)
        except Exception as e:
            print(f"扫描business_data目录失败: {str(e)}")
            return False

    def scan_func1(self) -> bool:
        """扫描 func1 目录并生成 target_pages.json"""
        try:
            func1_word_dir = self.file_manager.get_func1_dir('nostamped_word')
            word_files = self.file_processor.list_files(func1_word_dir, ['.docx', '.doc'])

            from GaiZhangYe.core.basic.word_processor import WordProcessor
            wp = WordProcessor()

            func1_data = {}
            for word_file in word_files:
                try:
                    page_count = wp.get_word_page_count(word_file)
                except Exception as e:
                    print(f"获取文件页数失败 {word_file.name}: {e}")
                    page_count = 0

                func1_data[word_file.stem] = {
                    'pages': [],
                    'total_pages': page_count
                }

            self.save_func1_data(func1_data)
            print(f"扫描到 {len(word_files)} 个Word文件，已生成func1默认数据")
            return True
        except Exception as e:
            print(f"扫描func1失败: {str(e)}")
            return False

    def scan_func2(self) -> bool:
        """扫描 func2 目录并生成 stamp_config.json"""
        try:
            func2_target_dir = self.file_manager.get_func2_dir('target_files')
            func2_image_dir = self.file_manager.get_func2_dir('images')

            target_files = self.file_processor.list_files(func2_target_dir, ['.docx'])
            image_files = self.file_processor.list_files(func2_image_dir, ['.png', '.jpg', '.jpeg'])

            # 对文件和图片进行排序 (使用Windows自然排序)
            sorted_target_files = sorted(target_files, key=lambda f: windows_natural_sort_key(f.name))
            sorted_image_files = sorted(image_files, key=lambda f: windows_natural_sort_key(f.name))

            func2_config = {}

            from GaiZhangYe.core.basic.word_processor import WordProcessor
            wp = WordProcessor()

            for i, target_file in enumerate(sorted_target_files):
                assigned_images = [sorted_image_files[i].name] if i < len(sorted_image_files) else []

                try:
                    total_pages = wp.get_word_page_count(target_file)
                    total_pages = total_pages if total_pages > 0 else 1
                except Exception as e:
                    print(f"获取文件页数失败 {target_file.name}: {e}")
                    total_pages = 1

                func2_config[target_file.name] = {
                    'total_pages': total_pages,
                    'images': assigned_images,
                    'positions': [{
                        'page': total_pages,
                        'x': 100,
                        'y': 100
                    } for _ in assigned_images]
                }

            func2_new_data = {
                "target_files": [f.name for f in sorted_target_files],
                "images": [f.name for f in sorted_image_files],
                "config": func2_config
            }

            self.save_func2_data(func2_new_data)
            print(f"扫描到 {len(target_files)} 个目标文件和 {len(image_files)} 个图片文件，已生成func2默认数据")
            return True
        except Exception as e:
            print(f"扫描func2失败: {str(e)}")
            return False
    
    def auto_generate_data(self) -> bool:
        """项目启动时自动生成数据文件"""
        try:
            # 检查并生成func1数据文件
            if not self.func1_data_file.exists():
                # 如果数据文件不存在，自动扫描生成
                self.scan_business_data()
            else:
                # 数据文件已存在，检查是否为空
                data = self.get_func1_data()
                if not data:
                    # 如果数据为空，重新扫描生成
                    self.scan_business_data()
            
            # 检查并生成func2数据文件
            if not self.func2_data_file.exists():
                self.scan_business_data()
            else:
                data = self.get_func2_data()
                if not data:
                    self.scan_business_data()
            
            return True
        except Exception as e:
            return False


# 单例模式
_data_service = None


def get_data_service() -> DataCommunicationService:
    """获取数据服务实例"""
    global _data_service
    if _data_service is None:
        _data_service = DataCommunicationService()
        # 应用启动时自动生成数据
        _data_service.auto_generate_data()
    return _data_service

