#!/usr/bin/env python3
"""
测试盖章页覆盖功能的使用示例
"""
from pathlib import Path
from typing import List
from dataclasses import dataclass
from GaiZhangYe.core.stamp_overlay import StampOverlayService
from GaiZhangYe.core.basic.pdf_processor import PdfProcessor

# 定义配置类
@dataclass
class UIConfig:
    """UI层传递的配置信息"""
    filename: str
    image_files: List[str]
    insert_positions: List[str]


def main():
    # 创建服务实例
    service = StampOverlayService()
    pdf_processor = PdfProcessor()

    # 配置参数
    target_word_dir = Path("GaiZhangYe/business_data/func2/TargetFiles")
    image_width = None  # 不缩放图片

    # 查看TargetFiles目录中的文件
    print("TargetFiles目录中的文件:")
    target_files = list(target_word_dir.iterdir())
    for file in target_files:
        print(f"  - {file.name}")
    print()

    # 查看Images目录中的文件
    print("Images目录中的文件:")
    images_dir = Path("GaiZhangYe/business_data/func2/Images")
    pdf_files = list(images_dir.glob("*.pdf"))

    for file in pdf_files:
        print(f"  - {file.name} (PDF文件)")

    # 将PDF转换为图片
    print("\n正在将PDF转换为图片...")
    image_files = []
    for pdf_file in pdf_files:
        extracted_images = pdf_processor.extract_images(pdf_file, images_dir)
        print(f"  - 从 {pdf_file.name} 提取了 {len(extracted_images)} 张图片")
        image_files.extend(extracted_images)

    print("\n转换后可用的图片:")
    for file in image_files:
        print(f"  - {file.name}")

    print()

    # 准备UI层配置
    # 这里我们使用转换后的第一张图片作为示例
    if image_files:
        configs = [
            # 文件1：插入1张图片，位置为last
            UIConfig(
                filename="第二部分1-3-5【项目组】26华远陆港公司债-履行普通注意义务的相关记录【需要签字】.docx",
                image_files=[str(image_files[0].name)],
                insert_positions=["last"]
            ),
            # 文件2：插入1张图片，位置为1
            UIConfig(
                filename="第二部分1-7【项目组】关于本次债券是否存在涉贿情况的自查文件.docx",
                image_files=[str(image_files[0].name)],
                insert_positions=["1"]
            ),
            # 文件3：插入1张图片，位置为last
            UIConfig(
                filename="第二部分1-9【项目组】公司债券项目尽职调查报告-25晋能电力公司债【最后更新】.docx",
                image_files=[str(image_files[0].name)],
                insert_positions=["last"]
            ),
            # 文件4：插入3张图片，位置分别为1, 2, 3
            UIConfig(
                filename="华泰联合证券有限责任公司关于晋能控股煤业集团有限公司关于2025年中期报告诉讼情况的受托管理事务临时报告-1215(1).docx",
                image_files=[str(image_files[0].name) for _ in range(3)],  # 使用同一图片插入3次
                insert_positions=["1", "2", "3"]
            )
        ]

        try:
            # 执行功能
            print("开始执行盖章页覆盖功能...")
            print()

            result_word_files = service.run(
                target_word_dir=target_word_dir,
                image_width=image_width,
                image_files=image_files,  # 这里传入的是实际存在的图片文件列表
                configs=configs
            )

            print("\n处理完成！生成的Word文件：")
            for file in result_word_files:
                print(f"  - {file}")

        except Exception as e:
            print(f"\n处理失败：{str(e)}")
            import traceback
            traceback.print_exc()
    else:
        print("没有可用的图片文件！")


if __name__ == "__main__":
    main()
