"""
基础功能测试
"""
import unittest
from pathlib import Path
from GaiZhangYe.utils.logger import get_logger


class TestBasicFunctionality(unittest.TestCase):
    """测试基础功能"""

    def test_logger_initialization(self):
        """测试日志器初始化"""
        logger = get_logger(__name__)
        self.assertIsNotNone(logger)
        self.assertEqual(logger.name, __name__)

    def test_project_structure(self):
        """测试项目目录结构"""
        project_root = Path(__file__).parent.parent

        # 检查核心目录是否存在
        self.assertTrue((project_root / "GaiZhangYe").exists())
        self.assertTrue((project_root / "GaiZhangYe" / "core").exists())
        self.assertTrue((project_root / "GaiZhangYe" / "utils").exists())
        self.assertTrue((project_root / "GaiZhangYe" / "web").exists())

        # 检查配置文件是否存在
        self.assertTrue((project_root / "pyproject.toml").exists())
        self.assertTrue((project_root / ".env.example").exists())


if __name__ == "__main__":
    unittest.main()
