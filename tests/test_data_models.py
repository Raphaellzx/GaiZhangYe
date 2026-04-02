"""
数据模型测试用例：测试项目中定义的业务数据模型
"""
import pytest
from pathlib import Path
from datetime import datetime
from GaiZhangYe.core.models.data_models import (
    Func1Data, PageConfig, Func2Data, StampConfig, ImageLocationType,
    ProcessResult, WordFileInfo, PdfFileInfo, ImageFileInfo, ProcessingContext
)
from GaiZhangYe.core.models.exceptions import BusinessError


class TestFunc1Data:
    """功能1数据模型测试"""

    def test_func1_data_creation(self):
        """测试Func1Data对象创建"""
        data = Func1Data(target_pages={"test.docx": [1, 2, 3]})
        assert isinstance(data, Func1Data)
        assert data.target_pages == {"test.docx": [1, 2, 3]}

    def test_func1_data_default(self):
        """测试Func1Data默认值"""
        data = Func1Data()
        assert isinstance(data, Func1Data)
        assert data.target_pages == {}

    def test_func1_data_validation(self):
        """测试Func1Data数据验证"""
        with pytest.raises(ValueError):
            Func1Data(target_pages="invalid")
        with pytest.raises(ValueError):
            Func1Data(target_pages={"test.docx": "invalid"})
        with pytest.raises(ValueError):
            Func1Data(target_pages={"": [1, 2]})


class TestFunc2Data:
    """功能2数据模型测试"""

    def test_func2_data_creation(self):
        """测试Func2Data对象创建"""
        configs = [
            StampConfig(
                filename="test1.docx",
                image_files=["stamp1.png", "stamp2.png"],
                positions=[1, 2],
                image_width=100
            ),
            StampConfig(
                filename="test2.docx",
                image_files=["stamp3.png"],
                positions=["last_page"],
                image_width=150
            )
        ]
        data = Func2Data(configs=configs)
        assert isinstance(data, Func2Data)
        assert len(data.configs) == 2
        assert all(isinstance(config, StampConfig) for config in data.configs)

    def test_func2_data_validation(self):
        """测试Func2Data数据验证"""
        with pytest.raises(ValueError):
            Func2Data(configs="invalid")


class TestStampConfig:
    """盖章配置数据模型测试"""

    def test_stamp_config_creation(self):
        """测试StampConfig对象创建"""
        config = StampConfig(
            filename="test.docx",
            image_files=["stamp.png"],
            positions=[1],
            image_width=100
        )
        assert isinstance(config, StampConfig)
        assert config.filename == "test.docx"
        assert config.image_files == ["stamp.png"]
        assert config.positions == [1]
        assert config.image_width == 100

    def test_stamp_config_validation(self):
        """测试StampConfig数据验证"""
        with pytest.raises(ValueError):
            StampConfig(filename="test.docx", image_files=[], positions=[1])
        with pytest.raises(ValueError):
            StampConfig(filename="test.docx", image_files=["stamp.png"], positions=[])
        with pytest.raises(ValueError):
            StampConfig(filename="", image_files=["stamp.png"], positions=[1])
        with pytest.raises(ValueError):
            StampConfig(filename="test.docx", image_files=["stamp.png"], positions=[1], image_width=0)
        with pytest.raises(ValueError):
            StampConfig(filename="test.docx", image_files=["stamp.png"], positions=[1], image_width=-10)


class TestImageLocationType:
    """图片位置类型测试"""

    def test_image_location_type(self):
        """测试ImageLocationType枚举"""
        assert ImageLocationType.PAGE_NUMBER == ImageLocationType(1)
        assert ImageLocationType.LAST_PAGE == ImageLocationType(2)
        assert ImageLocationType.PAGE_NUMBER.value == 1
        assert ImageLocationType.LAST_PAGE.value == 2


class TestProcessResult:
    """处理结果数据模型测试"""

    def test_process_result_creation(self):
        """测试ProcessResult对象创建"""
        result = ProcessResult(
            success=True,
            processed_files=[Path("file1.pdf"), Path("file2.pdf")],
            failed_files=[
                {"file": "file3.pdf", "error": "处理失败"},
                {"file": "file4.pdf", "error": "格式错误"}
            ],
            total_time=10.5,
            message="处理成功"
        )
        assert isinstance(result, ProcessResult)
        assert result.success
        assert len(result.processed_files) == 2
        assert len(result.failed_files) == 2
        assert result.total_time == 10.5
        assert result.message == "处理成功"

    def test_process_result_validation(self):
        """测试ProcessResult数据验证"""
        with pytest.raises(ValueError):
            ProcessResult(processed_files="invalid")
        with pytest.raises(ValueError):
            ProcessResult(failed_files="invalid")
        with pytest.raises(ValueError):
            ProcessResult(total_time=-1)


class TestFileInfoModels:
    """文件信息数据模型测试"""

    def test_word_file_info(self, tmp_path):
        """测试WordFileInfo"""
        test_file = tmp_path / "test.docx"
        test_file.touch()

        info = WordFileInfo(
            file_path=test_file,
            page_count=5,
            file_size=1024,
            last_modified=datetime.now()
        )

        assert isinstance(info, WordFileInfo)
        assert info.file_path == test_file
        assert info.page_count == 5
        assert info.file_size == 1024

    def test_pdf_file_info(self, tmp_path):
        """测试PdfFileInfo"""
        test_file = tmp_path / "test.pdf"
        test_file.touch()

        info = PdfFileInfo(
            file_path=test_file,
            page_count=3,
            file_size=2048,
            last_modified=datetime.now()
        )

        assert isinstance(info, PdfFileInfo)
        assert info.file_path == test_file
        assert info.page_count == 3
        assert info.file_size == 2048

    def test_image_file_info(self, tmp_path):
        """测试ImageFileInfo"""
        test_file = tmp_path / "test.png"
        test_file.touch()

        info = ImageFileInfo(
            file_path=test_file,
            width=100,
            height=100,
            file_size=512,
            last_modified=datetime.now()
        )

        assert isinstance(info, ImageFileInfo)
        assert info.file_path == test_file
        assert info.width == 100
        assert info.height == 100
        assert info.file_size == 512

    def test_file_info_validation(self, tmp_path):
        """测试文件信息数据验证"""
        test_file = tmp_path / "test.docx"
        test_file.touch()

        with pytest.raises(ValueError):
            WordFileInfo(file_path="invalid")
        with pytest.raises(ValueError):
            WordFileInfo(file_path=test_file, page_count=-1)
        with pytest.raises(FileNotFoundError):
            WordFileInfo(file_path=Path("nonexistent_file.docx"))


class TestProcessingContext:
    """处理上下文数据模型测试"""

    def test_processing_context_creation(self):
        """测试ProcessingContext对象创建"""
        context = ProcessingContext(
            task_id="test_task_123",
            start_time=datetime.now(),
            parameters={"param1": "value1", "param2": 100},
            environment={"os": "Windows", "python_version": "3.10"}
        )

        assert isinstance(context, ProcessingContext)
        assert context.task_id == "test_task_123"
        assert isinstance(context.start_time, datetime)
        assert context.parameters == {"param1": "value1", "param2": 100}
        assert context.environment == {"os": "Windows", "python_version": "3.10"}

    def test_processing_context_validation(self):
        """测试ProcessingContext数据验证"""
        with pytest.raises(ValueError):
            ProcessingContext(task_id="", start_time=datetime.now())
        with pytest.raises(ValueError):
            ProcessingContext(task_id="test", start_time="invalid")


class TestIntegration:
    """数据模型集成测试"""

    def test_full_workflow(self, tmp_path):
        """测试完整工作流程数据模型集成"""
        # 创建功能1数据
        func1_data = Func1Data(target_pages={"doc1.docx": [1, 3], "doc2.docx": [2]})

        # 创建功能2数据
        configs = [
            StampConfig(
                filename="doc1.docx",
                image_files=["stamp1.png", "stamp2.png"],
                positions=[1, "last_page"],
                image_width=100
            ),
            StampConfig(
                filename="doc2.docx",
                image_files=["stamp3.png"],
                positions=[2],
                image_width=150
            )
        ]
        func2_data = Func2Data(configs=configs)

        # 创建处理结果
        result = ProcessResult(
            success=True,
            processed_files=[tmp_path / "doc1_processed.docx", tmp_path / "doc2_processed.docx"],
            failed_files=[],
            total_time=15.2,
            message="所有文件处理成功"
        )

        # 验证所有数据模型是否正常工作
        assert isinstance(func1_data, Func1Data)
        assert isinstance(func2_data, Func2Data)
        assert isinstance(result, ProcessResult)
        assert len(func2_data.configs) == 2
        assert all(isinstance(config, StampConfig) for config in func2_data.configs)
