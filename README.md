GaiZhangYe 项目说明文档

1\. 项目概述

1.1 项目名称

GaiZhangYe（盖章页处理工具）

1.2 项目定位

一款基于 Python 开发的桌面端工具，专注解决文档盖章页处理的核心需求，实现 "盖章页准备""盖章页覆盖""批量 Word 转 PDF" 三大核心功能，适配 Windows 环境，依托 pywin32、pymupdf、pillow 实现文档 / 图片处理，通过 tkinter 提供友好的图形交互界面。

1.3 核心功能

  -----------------------------------------------------------------------
  功能模块           功能描述
  ------------------ ----------------------------------------------------
  盖章页准备         将未盖章 Word 文件转换为 PDF，提取指定页码生成待盖章的 PDF 页面文件

  盖章页覆盖         从盖章后的 PDF / 图片中提取内容，缩放后插入到目标 Word 文件，生成最终 Word/PDF

  批量 Word 转 PDF   批量处理指定目录下的 Word 文件，一键转换为 PDF 格式
  -----------------------------------------------------------------------

1.4 技术栈

核心依赖：pywin32（Word 操作）、pymupdf（PDF 处理）、pillow（图片处理）、tkinter（UI 界面）

项目管理：uv（依赖管理 / 虚拟环境）、pydantic-settings（配置加载）

开发规范：Pytest（单元测试）、Ruff（代码规范）

2\. 环境要求

2.1 系统环境

操作系统：Windows 10/11（64 位）

依赖软件：Microsoft Word（2016 及以上版本，pywin32 需依赖本地 Word 客户端）

2.2 Python 环境

Python 版本：3.11.9

虚拟环境：uv 内置（无需手动创建）

3\. 安装指南

3.1 源码安装

步骤 1：克隆 / 下载项目

bash 运行

\# 克隆仓库（若使用Git）
git clone \<项目仓库地址\>
cd GaiZhangYe

步骤 2：安装 uv（依赖管理工具）

bash 运行

\# Windows PowerShell 执行
powershell -c "irm https://astral.sh/uv/install.ps1 | iex"

步骤 3：初始化项目并安装依赖

bash 运行

\# 初始化项目（指定Python版本）
uv init GaiZhangYe --python 3.11
\# 安装生产依赖（核心功能）
uv install
\# 可选：安装开发依赖（测试/代码规范）
uv install --dev

3.2 环境变量配置（可选）

项目根目录创建 .env 文件（参考 .env.example），支持自定义配置：

ini

\# 日志级别：DEBUG/INFO/WARNING/ERROR（默认 INFO）
LOG_LEVEL=INFO
\# 业务数据目录自定义路径（默认项目根/business_data）
BUSINESS_DATA_ROOT=D:/GaiZhangYe/business_data

4\. 目录结构

4.1 整体结构

plaintext
GaiZhangYe/ \# 项目根目录
├── pyproject.toml \# 依赖配置/项目元信息
├── uv.lock \# 依赖锁文件（环境一致性）
├── .env \# 环境变量（不上Git）
├── .env.example \# 环境变量示例（上Git）
├── .gitignore \# Git忽略规则
├── README.md \# 项目说明（本文档）
├── docs/ \# 扩展文档（架构/流程）
├── tests/ \# 测试目录（镜像源码结构）
├── GaiZhangYe/ \# 核心源码包
│   ├── core/ \# 核心业务层（纯逻辑，无UI）
│   │   ├── file_processor.py \# 通用文件处理（校验/遍历）
│   │   ├── file_manager.py \# 业务目录管理（创建/路径映射）
│   │   ├── word_processor.py \# Word专属处理（转PDF/插入图片）
│   │   ├── pdf_processor.py \# PDF专属处理（提取页面/图片）
│   │   ├── image_processor.py \# 图片专属处理（缩放/格式转换）
│   │   ├── models/ \# 数据模型（配置/异常）
│   │   └── services/ \# 业务服务（整合核心模块实现功能）
│   ├── ui/ \# UI层（tkinter，仅交互）
│   ├── utils/ \# 通用工具（日志/配置）
│   └── entrypoints/ \# 启动入口
├── business_data/ \# 业务数据目录（自动创建）
│   ├── func1/ \# 功能1：盖章页准备
│   │   ├── Nostamped_Word/ \# 输入：未盖章Word文件
│   │   ├── Nostamped_PDF/ \# 输出：Word转PDF文件
│   │   └── Stamped_Pages/ \# 输出：待盖章PDF页面
│   └── func2/ \# 功能2：盖章页覆盖
│       ├── Images/ \# 输入：盖章图片/PDF提取的图片
│       ├── TargetFiles/ \# 输入：目标Word文件
│       ├── Result_Word/ \# 输出：插入图片后的Word
│       └── Result_PDF/ \# 输出：Result_Word转PDF文件
├── logs/ \# 日志目录（自动生成）
└── .venv/ \# uv虚拟环境（自动创建）

4.2 核心目录说明

  ---------------------------------------------------------------------------
  目录 / 文件            核心作用
  ---------------------- ----------------------------------------------------
  core/                  核心业务逻辑层，按 "文件 / Word/PDF/ 图片" 维度拆分，无 UI 依赖
  core/services/         业务服务层，整合核心模块实现三大功能，是 UI 层调用的入口
  ui/                    纯交互层，仅负责界面展示、用户输入接收，调用 services 层实现业务逻辑
  business_data/         业务数据目录，严格区分功能 1/2 的输入输出，避免文件混乱
  utils/                 全局工具，包含日志配置、环境变量加载，所有模块复用
  ---------------------------------------------------------------------------

5\. 使用指南

5.1 启动项目

bash 运行

\# 启动图形界面（推荐）
uv run GaiZhangYe/entrypoints/gui.py

5.2 功能 1：盖章页准备

操作步骤

启动工具后，切换到 "1. 准备盖章页" 标签页；
（可选）自定义 "未盖章 Word 目录"（默认 business_data/func1/Nostamped_Word）；
输入需要提取的 PDF 页码（如 1,3,5，逗号分隔）；
点击 "开始准备"，工具自动完成：
将 Word 文件转换为 PDF 到 Nostamped_PDF；
提取指定页码生成 PDF 到 Stamped_Pages。

注意事项

仅支持 .docx/.doc 格式的 Word 文件；
页码需为正整数，且不超过 PDF 总页数。

5.3 功能 2：盖章页覆盖

操作步骤

切换到 "2. 盖章页覆盖" 标签页；
选择 "盖章 PDF / 图片路径"（支持 .pdf/.png/.jpg）；
（可选）设置图片缩放宽度（默认 800px，等比缩放）；
（可选）自定义 "目标 Word 目录"（默认 business_data/func2/TargetFiles）；
点击 "开始覆盖"，工具自动完成：
从 PDF / 图片提取内容到 Images；
缩放图片后插入到目标 Word 文件，生成到 Result_Word；
将 Result_Word 转换为 PDF 到 Result_PDF。

注意事项

图片格式推荐 PNG/JPG，PDF 需为可提取图片的格式；
目标 Word 文件建议提前清理冗余内容，避免插入位置异常。

5.4 功能 3：批量 Word 转 PDF

操作步骤

切换到 "3. 批量 Word 转 PDF" 标签页；
选择 "输入目录"（存放待转换的 Word 文件）；
选择 "输出目录"（存放生成的 PDF 文件）；
点击 "开始批量转换"，工具自动遍历输入目录的 Word 文件并转换。

注意事项

支持批量处理 .docx/.doc 格式，忽略非 Word 文件；
输出目录若不存在，工具会自动创建。

6\. 开发指南

6.1 代码开发顺序

参考 "从底层到上层" 的原则，推荐顺序：
基础搭建（目录 / 配置 / 日志）→ 2. Core 核心模块 → 3. 业务服务层 → 4. UI 层 → 5. 测试 & 优化

6.2 核心模块分工

  --------------------------------------------------------------------------------------------------
  模块文件                   职责范围
  -------------------------- -----------------------------------------------------------------------
  file_processor.py          通用文件操作：校验文件存在 / 类型、遍历目录文件、删除临时文件
  file_manager.py            业务目录管理：创建 business_data 及子目录、清理过期文件、返回目录路径
  word_processor.py          Word 专属操作：单文件 / 批量转 PDF、插入图片到 Word
  pdf_processor.py           PDF 专属操作：提取指定页面、提取 PDF 内图片、PDF 合法性校验
  image_processor.py         图片专属操作：等比缩放、格式转换、图片合法性校验
  --------------------------------------------------------------------------------------------------

6.3 测试

bash 运行

\# 运行所有测试
uv run pytest tests/ -v
\# 运行指定模块测试（如Word处理器）
uv run pytest tests/test_core/test_word_processor.py -v

6.4 代码规范

bash 运行

\# 代码格式检查/修复
uv run ruff check GaiZhangYe/ --fix

7\. 常见问题

7.1 启动报错 "找不到 Word 应用"

原因：本地未安装 Microsoft Word，或 pywin32 与 Word 版本不兼容；
解决：安装 Word 2016+，重新安装 pywin32（uv add pywin32>=306）。

7.2 Word 转 PDF 无输出

原因：Word 文件损坏，或权限不足；
解决：检查 Word 文件是否能正常打开，确保 business_data 目录有读写权限。

7.3 图片插入 Word 位置异常

原因：Word 文档格式特殊（如分栏、表格），插入位置默认在文档末尾；
解决：修改 word_processor.py 中 insert_image_to_word 方法的插入位置逻辑。

7.4 日志无输出

原因：日志级别配置过高，或 logs 目录无权限；
解决：修改 .env 中 LOG_LEVEL=DEBUG，检查 logs 目录权限。