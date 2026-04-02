```
# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.
```

## 项目概述

盖章页处理工具是一个Python应用程序，专为批量处理Word文档和PDF文件中的盖章页而设计。它提供了准备盖章页、盖章页覆盖以及批量Word转PDF等核心功能，支持命令行界面(CLI)和Web界面两种操作方式。

## 核心功能

### 主要功能模块
- **准备盖章页**: 为Word文档自动生成符合规范的盖章页
- **盖章页覆盖**: 将盖章页精确覆盖到PDF文件的指定位置
- **批量Word转PDF**: 将多个Word文档批量转换为PDF格式

### 操作方式
- **CLI命令行**: 支持批量处理和自动化脚本
- **Web界面**: 提供直观的图形界面操作

## 项目结构

```
GaiZhangYe/
├── business_data/       # 业务数据
├── core/               # 核心功能
│   ├── basic/          # 基础处理模块
│   ├── entrypoints/    # 入口文件
│   ├── models/         # 数据模型和配置
│   ├── stamp_prepare.py # 准备盖章页功能
│   ├── stamp_overlay.py # 盖章页覆盖功能
│   └── batch_convert.py # 批量Word转PDF功能
├── utils/              # 工具函数
├── web/                # Web界面
│   ├── templates/      # HTML模板
│   │   ├── index.html  # 首页模板
│   │   └── pages/      # 功能页面模板
│   ├── static/         # 静态资源
│   │   └── css/        # 样式文件
│   │       └── theme.css # 主题样式文件
│   └── routes/         # 路由和API
│       ├── api.py      # API接口实现
│       └── pages.py    # 页面路由实现
└── tests/              # 测试文件
```

## 开发环境

### 环境要求
- Python 3.10.19或更高版本
- Windows系统(Word处理依赖pywin32)

### 依赖管理
项目使用uv作为推荐的包管理器，同时支持传统的pip方式。

```bash
# 使用uv创建虚拟环境和安装依赖
uv sync -p 3.10.19

# 激活虚拟环境(Windows)
.venv\Scripts\activate

# 传统方式安装依赖
pip install -r requirements.txt
```

## 开发流程

### 运行应用

#### 命令行界面
```bash
# 准备盖章页
python -m GaiZhangYe.core.entrypoints.cli_start --prepare

# 盖章页覆盖
python -m GaiZhangYe.core.entrypoints.cli_start --cover

# 批量Word转PDF
python -m GaiZhangYe.core.entrypoints.cli_start --convert
```

#### Web界面
```bash
# 方式1：使用入口文件启动（推荐）
python -m GaiZhangYe.core.entrypoints.start_service

# 方式2：直接启动Flask应用
python -m GaiZhangYe.web.app

# 访问 http://localhost:5001
```

### 运行测试

```bash
# 运行所有测试
pytest

# 运行特定测试文件
pytest tests/test_xxx.py

# 运行测试并显示详细输出
pytest -v

# 测试API接口是否正常工作
python test_api.py
```

### 代码规范检查

```bash
# 检查代码规范
ruff check .

# 自动修复代码规范问题
ruff fix .
```

## 配置

配置文件采用.env格式，参考.env.example创建.env文件：

```bash
cp .env.example .env
```

## 构建

使用PyInstaller构建可执行文件：

```bash
pyinstaller GaiZhangYe.spec
```

构建完成后，可执行文件将生成在dist/目录下。

## 核心技术栈

- **Word处理**: pywin32 (Windows系统)
- **PDF处理**: PyMuPDF (pymupdf)
- **图片处理**: Pillow
- **Web框架**: Flask
- **配置管理**: Pydantic Settings + python-dotenv
- **测试**: pytest
- **代码规范**: ruff

## 关键文件说明

| 文件路径 | 功能说明 |
|---------|---------|
| GaiZhangYe/core/basic/file_manager.py | 文件管理类，提供文件和目录操作的统一接口 |
| GaiZhangYe/core/basic/word_processor.py | Word文档处理基础功能，封装pywin32的Word操作 |
| GaiZhangYe/core/basic/pdf_processor.py | PDF文件处理基础功能，使用PyMuPDF库处理PDF文件 |
| GaiZhangYe/core/basic/image_processor.py | 图片处理基础功能，使用Pillow库处理图片 |
| GaiZhangYe/core/stamp_prepare.py | 准备盖章页功能实现 |
| GaiZhangYe/core/stamp_overlay.py | 盖章页覆盖功能实现 |
| GaiZhangYe/core/batch_convert.py | 批量Word转PDF功能实现 |
| GaiZhangYe/web/app.py | Web应用入口和配置 |
| GaiZhangYe/core/entrypoints/start_service.py | HTML可视化服务入口文件，自动清除旧进程并启动Flask应用 |
| GaiZhangYe/web/routes/api.py | API接口实现 |
| GaiZhangYe/utils/config.py | 全局配置模型，从.env文件加载配置，支持业务目录和Web服务配置 |
| GaiZhangYe/utils/logger.py | 日志工具函数，提供日志配置和记录功能 |
| GaiZhangYe/web/routes/pages.py | 页面路由实现 |
| tests/test_api.py | API接口测试文件，直接在Flask应用上下文中测试API，包含路由检查和接口测试 |
| tests/test_stamp_prepare.py | 准备盖章页功能测试文件，验证盖章页准备过程的正确性 |
| tests/test_stamp_overlay.py | 盖章页覆盖功能测试文件，验证盖章页覆盖过程的正确性 |
| GaiZhangYe/core/models/data_models.py | 数据模型定义文件，包含Func1Data、Func2Data等数据模型类 |
| GaiZhangYe/core/models/exceptions.py | 异常定义文件，包含FileProcessError等自定义异常类 |
| tests/test_data_models.py | 数据模型测试文件，验证数据模型的正确性和完整性 |
