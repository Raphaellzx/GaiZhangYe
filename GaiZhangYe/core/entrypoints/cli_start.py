#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
命令行启动版本 - 直接调用核心功能
"""

import sys
import logging
import argparse

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def main():
    parser = argparse.ArgumentParser(description='盖章页工具命令行版本')
    parser.add_argument('-f', '--function', choices=['1', '2'], required=True,
                       help='选择功能: 1-准备盖章页, 2-盖章页覆盖')
    
    args = parser.parse_args()
    
    if args.function == '1':
        from GaiZhangYe.core.stamp_prepare import StampPrepareService
        service = StampPrepareService()
        result = service.run()
        
        if result:
            logger.info(f'功能1执行完成，生成文件: {len(result)}个')
            for file in result:
                logger.info(f'  - {file}')
        else:
            logger.warning('功能1执行完成，没有生成任何文件')
            
    elif args.function == '2':
        from GaiZhangYe.core.stamp_overlay import StampOverlayService
        service = StampOverlayService()
        word_result, pdf_result = service.run()
        
        if word_result or pdf_result:
            if word_result:
                logger.info(f'功能2执行完成，生成Word文件: {len(word_result)}个')
            if pdf_result:
                logger.info(f'功能2执行完成，生成PDF文件: {len(pdf_result)}个')
        else:
            logger.warning('功能2执行完成，没有生成任何文件')

if __name__ == '__main__':
    main()
