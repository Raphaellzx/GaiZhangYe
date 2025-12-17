#!/usr/bin/env python3
"""
UI层与盖章页覆盖服务交互的示例
"""
from pathlib import Path
from dataclasses import dataclass
from typing import List, Optional

# 假设这是UI层定义的配置类
@dataclass
class UIBatchConfig:
    """UI层批量处理配置"""
    target_word_dir: str
    image_width: Optional[int]
    image_files: List[str]  # UI层选择的图片文件路径列表
    file_configs: List  # 每个文件的配置


# 定义单个文件的配置
@dataclass
class UIFileConfig:
    """单个文件的UI配置"""
    filename: str
    image_files: List[str]  # 该文件需要插入的图片文件路径或文件名
    insert_positions: List[str]  # 对应的插入位置


def ui_to_service_config(ui_config: UIBatchConfig):
    """将UI层配置转换为服务需要的格式"""
    from GaiZhangYe.core.stamp_overlay import StampOverlayService

    # 创建服务实例
    service = StampOverlayService()

    try:
        # 将UI层的配置转换为服务需要的类型
        target_dir = Path(ui_config.target_word_dir)
        image_files = [Path(img) for img in ui_config.image_files]

        # 准备文件配置
        configs = ui_config.file_configs

        # 调用服务
        result_word_files = service.run(
            target_word_dir=target_dir,
            image_width=ui_config.image_width,
            image_files=image_files,
            configs=configs
        )

        # 返回处理结果给UI层
        return {
            "success": True,
            "message": f"成功处理{len(result_word_files)}个文件",
            "result_files": [str(f) for f in result_word_files]
        }

    except Exception as e:
        return {
            "success": False,
            "message": f"处理失败: {str(e)}",
            "result_files": []
        }


def demo():
    """演示如何使用UI配置"""

    # 模拟UI层收集的配置信息
    ui_config = UIBatchConfig(
        target_word_dir="GaiZhangYe/business_data/func2/TargetFiles",
        image_width=80,  # 缩放宽度为80
        image_files=[
            "GaiZhangYe/business_data/func2/Images/盖章页文件.pdf"  # UI层选择的图片文件
        ],
        file_configs=[
            # 文件1配置
            UIFileConfig(
                filename="第二部分1-3-5【项目组】26华远陆港公司债-履行普通注意义务的相关记录【需要签字】.docx",
                image_files=["盖章页文件.pdf"],  # 可以是文件名
                insert_positions=["last"]
            ),
            # 文件2配置
            UIFileConfig(
                filename="第二部分1-7【项目组】关于本次债券是否存在涉贿情况的自查文件.docx",
                image_files=["盖章页文件.pdf"],  # 可以是完整路径
                insert_positions=["1"]
            ),
            # 文件3配置
            UIFileConfig(
                filename="第二部分1-9【项目组】公司债券项目尽职调查报告-25晋能电力公司债【最后更新】.docx",
                image_files=["盖章页文件.pdf"],  # 可以是文件名
                insert_positions=["last"]
            ),
            # 文件4配置
            UIFileConfig(
                filename="华泰联合证券有限责任公司关于晋能控股煤业集团有限公司关于2025年中期报告诉讼情况的受托管理事务临时报告-1215(1).docx",
                image_files=["盖章页文件.pdf", "盖章页文件.pdf", "盖章页文件.pdf"],  # 多个图片
                insert_positions=["1", "2", "3"]  # 多个位置
            )
        ]
    )

    # 调用服务
    result = ui_to_service_config(ui_config)

    # 显示结果
    print("处理结果：")
    print(f"成功：{result['success']}")
    print(f"消息：{result['message']}")
    if result['result_files']:
        print(f"生成文件：")
        for file in result['result_files']:
            print(f"  - {file}")


if __name__ == "__main__":
    demo()
