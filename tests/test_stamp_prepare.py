#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
功能1：准备盖章页服务的测试文件
使用 tests/test1 目录中的数据模拟输入
"""
import pytest
from pathlib import Path
from GaiZhangYe.core.stamp_prepare import StampPrepareService
from GaiZhangYe.core.data_communication import get_data_service
from GaiZhangYe.core.models.data_models import Func1Data
from GaiZhangYe.core.models.exceptions import BusinessError


class TestStampPrepareService:
    """功能1：准备盖章页服务的测试类"""

    @classmethod
    def setup_class(cls):
        """在测试类开始前执行"""
        cls.test_dir = Path("tests/test1")
        cls.word_dir = cls.test_dir / "word"
        cls.result_pdf_dir = cls.test_dir / "result_pdf"
        cls.service = StampPrepareService()

    def test_initialization(self):
        """测试服务初始化"""
        assert self.service is not None

    def test_stamp_prepare_with_valid_data(self):
        """使用有效的测试数据执行功能1"""
        # 获取目录中的实际文件名
        actual_files = list(self.word_dir.iterdir())
        actual_filenames = [f.name for f in actual_files if f.suffix in [".docx", ".doc"]]

        # 设置测试数据
        data_service = get_data_service()
        test_data = Func1Data(target_pages={})

        for filename in actual_filenames:
            test_data.target_pages[filename.rstrip(".docx").rstrip(".doc")] = [1]

        data_service.save_func1_data(test_data)

        try:
            # 执行功能1
            result = self.service.run(
                word_dir=self.word_dir,
                output_dir=self.result_pdf_dir
            )

            # 验证结果
            assert len(result) > 0
            for pdf_path in result:
                assert pdf_path.exists()
                assert pdf_path.suffix == ".pdf"

            print(f"测试成功，生成了 {len(result)} 个PDF文件")
        finally:
            # 清理测试数据
            data_service.save_func1_data(Func1Data())

    def test_stamp_prepare_with_invalid_data(self):
        """测试无效数据的处理"""
        data_service = get_data_service()
        data_service.save_func1_data(Func1Data(target_pages={}))

        with pytest.raises(BusinessError):
            self.service.run(
                word_dir=self.word_dir,
                output_dir=self.result_pdf_dir
            )

    def test_stamp_prepare_with_missing_files(self):
        """测试目录中没有Word文件的情况"""
        data_service = get_data_service()
        data_service.save_func1_data(Func1Data(target_pages={"nonexistent.docx": [1]}))

        with pytest.raises(BusinessError):
            self.service.run(
                word_dir=self.result_pdf_dir,  # 这个目录中没有Word文件
                output_dir=self.result_pdf_dir
            )

    def test_cleanup(self):
        """测试资源清理"""
        data_service = get_data_service()
        data_service.save_func1_data(Func1Data())
