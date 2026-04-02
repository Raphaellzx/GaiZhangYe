#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
功能2：盖章页覆盖服务的测试文件
使用 tests/test2 目录中的数据模拟输入
"""
import pytest
from pathlib import Path
from GaiZhangYe.core.stamp_overlay import StampOverlayService
from GaiZhangYe.core.data_communication import get_data_service
from GaiZhangYe.core.models.data_models import Func2Data, StampConfig
from GaiZhangYe.core.models.exceptions import BusinessError


class TestStampOverlayService:
    """功能2：盖章页覆盖服务的测试类"""

    @classmethod
    def setup_class(cls):
        """在测试类开始前执行"""
        cls.test_dir = Path("tests/test2")
        cls.word_dir = cls.test_dir / "word"
        cls.result_word_dir = cls.test_dir / "盖章页word"
        cls.result_pdf_dir = cls.test_dir / "盖章页pdf"
        cls.stamp_pdf = cls.test_dir / "盖章页.pdf"
        cls.service = StampOverlayService()

        # 从PDF中提取图片
        from GaiZhangYe.core.basic.pdf_processor import PdfProcessor
        cls.pdf_processor = PdfProcessor()
        cls.images_dir = cls.test_dir / "extracted_images"
        cls.images_dir.mkdir(exist_ok=True)
        cls.extracted_images = cls.pdf_processor.extract_images(cls.stamp_pdf, cls.images_dir)

    def test_initialization(self):
        """测试服务初始化"""
        assert self.service is not None

    def test_stamp_overlay_with_valid_config(self):
        """使用有效的配置数据执行功能2"""
        # 获取目录中的实际文件名
        actual_files = list(self.word_dir.iterdir())
        actual_filenames = [f.name for f in actual_files if f.suffix in [".docx", ".doc"]]

        # 创建测试配置
        data_service = get_data_service()
        test_config = Func2Data(configs=[])

        for filename in actual_filenames:
            if self.extracted_images:
                # 确保图片数量和位置数量一致
                positions = [1] * len(self.extracted_images)
                test_config.configs.append(
                    StampConfig(
                        filename=filename,
                        image_files=[str(img_path) for img_path in self.extracted_images],
                        positions=positions
                    )
                )

        data_service.save_func2_data(test_config)

        try:
            # 执行功能2
            result = self.service.run(
                target_word_dir=self.word_dir,
                result_word_dir=self.result_word_dir,
                result_pdf_dir=self.result_pdf_dir,
                image_width=100,
                configs=test_config
            )

            # 验证结果
            assert len(result) > 0
            for word_path in result:
                assert word_path.exists()
                assert word_path.suffix == ".docx"
                # 验证对应的PDF文件是否生成
                pdf_path = self.result_pdf_dir / f"{word_path.stem}.pdf"
                assert pdf_path.exists()

            print(f"测试成功，生成了 {len(result)} 个Word文件")
        finally:
            # 清理测试数据
            data_service.save_func2_data(Func2Data())

    def test_stamp_overlay_with_default_mode(self):
        """测试默认模式（无配置）执行功能2"""
        data_service = get_data_service()
        data_service.save_func2_data(Func2Data())

        if not self.extracted_images:
            pytest.skip("未从PDF中提取到图片，跳过测试")

        result = self.service.run(
            target_word_dir=self.word_dir,
            result_word_dir=self.result_word_dir,
            result_pdf_dir=self.result_pdf_dir,
            image_files=self.extracted_images,
            image_width=100
        )

        assert len(result) > 0
        for word_path in result:
            assert word_path.exists()
            assert word_path.suffix == ".docx"

    def test_stamp_overlay_with_invalid_image(self):
        """测试使用无效图片路径的处理"""
        invalid_image_path = self.test_dir / "nonexistent_image.png"
        data_service = get_data_service()
        data_service.save_func2_data(Func2Data())

        with pytest.raises(BusinessError):
            self.service.run(
                target_word_dir=self.word_dir,
                result_word_dir=self.result_word_dir,
                result_pdf_dir=self.result_pdf_dir,
                image_files=[invalid_image_path],
                image_width=100
            )

    def test_stamp_overlay_with_missing_word_files(self):
        """测试Word文件不存在的情况"""
        data_service = get_data_service()
        test_config = Func2Data(configs=[
            StampConfig(
                filename="nonexistent.docx",
                image_files=[str(img_path) for img_path in self.extracted_images],
                positions=[1]
            )
        ])
        data_service.save_func2_data(test_config)

        result = self.service.run(
            target_word_dir=self.word_dir,
            result_word_dir=self.result_word_dir,
            result_pdf_dir=self.result_pdf_dir,
            image_width=100,
            configs=test_config
        )

        assert len(result) == 0

    def test_cleanup(self):
        """测试资源清理"""
        data_service = get_data_service()
        data_service.save_func2_data(Func2Data())
