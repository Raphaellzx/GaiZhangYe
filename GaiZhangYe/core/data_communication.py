#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
core层数据沟通模块
实现前后端通过文件进行数据交换的功能
"""

import json
from pathlib import Path
from typing import Dict, Any, Optional
from GaiZhangYe.core.basic.file_manager import FileManager
from GaiZhangYe.core.basic.file_processor import FileProcessor


class DataCommunicationService:
    """数据沟通服务类"""
    
    def __init__(self):
        # 定义数据文件路径
        self.func1_data_file = Path(__file__).parent.parent / 'business_data' / 'func1' / '.temp' / 'target_pages.json'
        self.func2_data_file = Path(__file__).parent.parent / 'business_data' / 'func2' / '.temp' / 'stamp_config.json'
        
        # 创建文件管理器和处理器实例
        self.file_manager = FileManager()
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
        """扫描business_data目录并生成默认数据"""
        try:
            # 扫描func1目录，生成默认target_pages
            func1_word_dir = self.file_manager.get_func1_dir('nostamped_word')
            word_files = self.file_processor.list_files(func1_word_dir, ['.docx', '.doc'])

            from GaiZhangYe.core.basic.word_processor import WordProcessor
            wp = WordProcessor()

            func1_data = {}
            for word_file in word_files:
                # 默认提取所有页面
                # 获取Word文件的实际页数
                try:
                    page_count = wp.get_word_page_count(word_file)
                except Exception as e:
                    print(f"获取文件页数失败 {word_file.name}: {e}")
                    page_count = 0

                func1_data[word_file.stem] = {
                    'pages': [],  # 选择的页面
                    'total_pages': page_count  # 总页数
                }
            
            # 保存func1数据
            self.save_func1_data(func1_data)
            print(f"扫描到 {len(word_files)} 个Word文件，已生成func1默认数据")
            
            # 扫描func2目录，生成默认stamp_config
            func2_target_dir = self.file_manager.get_func2_dir('target_files')
            func2_image_dir = self.file_manager.get_func2_dir('images')
            
            target_files = self.file_processor.list_files(func2_target_dir, ['.docx'])
            image_files = self.file_processor.list_files(func2_image_dir, ['.png', '.jpg', '.jpeg'])
            
            # 使用Windows资源管理器风格的排序函数
            def natural_sort_key(s):
                """自然排序键生成器，模拟Windows资源管理器排序"""
                import re
                return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)]

            # 对文件和图片进行排序
            sorted_target_files = sorted(target_files, key=lambda f: natural_sort_key(f.name))
            sorted_image_files = sorted(image_files, key=lambda f: natural_sort_key(f.name))

            func2_data = {}
            for i, target_file in enumerate(sorted_target_files):
                # 为每个文档分配对应的图片（按顺序一一对应）
                assigned_images = [sorted_image_files[i].name] if i < len(sorted_image_files) else []

                # 默认配置：每个文件对应一张图片，插入到最后一页
                func2_data[target_file.name] = {
                    'images': assigned_images,  # 按顺序一一对应
                    'positions': [{
                        'page': 'last_page',  # 默认最后一页
                        'x': 100,
                        'y': 100
                    } for _ in assigned_images]  # 每个图片对应一个位置
                }
            
            # 保存func2数据
            self.save_func2_data(func2_data)
            print(f"扫描到 {len(target_files)} 个目标文件和 {len(image_files)} 个图片文件，已生成func2默认数据")
            
            return True
        except Exception as e:
            print(f"扫描business_data目录失败: {str(e)}")
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
