# 盖章页处理工具

盖章页处理工具是一个Python应用程序，专为批量处理Word文档和PDF文件中的盖章页而设计。它提供了准备盖章页、盖章页覆盖以及批量Word转PDF等核心功能，支持命令行界面(CLI)和Web界面两种操作方式。

## 功能特性

### 📁 核心功能

- **准备盖章页**: 为Word文档自动生成符合规范的盖章页
- **盖章页覆盖**: 将盖章页精确覆盖到PDF文件的指定位置
- **批量Word转PDF**: 将多个Word文档批量转换为PDF格式

### 🎮 操作方式

- **CLI命令行**: 支持批量处理和自动化脚本
- **Web界面**: 提供直观的图形界面操作

## 安装

### 环境要求
- Python 3.10.19或更高版本
- Windows系统(Word处理依赖pywin32)

### 版本管理(推荐)
如果您需要管理多个Python版本，推荐使用pyenv：

```bash
# 安装pyenv (Windows: https://github.com/pyenv-win/pyenv-win)
# Linux/macOS: https://github.com/pyenv/pyenv

# 安装Python 3.10.19
pyenv install 3.10.19

# 在当前项目目录使用Python 3.10.19
pyenv local 3.10.19
```

pyenv会自动使用项目目录中的.python-version文件(已包含)来管理版本。

### 安装步骤

1. 克隆项目
   ```bash
   git clone <repository-url>
   cd GaiZhangYe
   ```

#### 使用uv创建虚拟环境和安装依赖(推荐)
uv是本项目推荐的现代化Python包和虚拟环境管理器，它会自动处理Python版本依赖：

```bash
# 确保已安装uv
pip install uv

# uv会自动：
# 1. 读取.pyenv-version或pyproject.toml中的requires-python配置
# 2. 创建符合要求的虚拟环境
# 3. 安装所有依赖
uv install
```

#### 激活虚拟环境
```bash
# Windows激活
.venv\Scripts\activate
# Linux/macOS激活
source .venv/bin/activate
```

#### 传统方式(可选)
```bash
# Windows
py -3.10 -m venv .venv && .venv\Scripts\activate && pip install -r requirements.txt
# Linux/macOS
python3.10 -m venv .venv && source .venv/bin/activate && pip install -r requirements.txt
```

## 使用

### 1. CLI命令行

#### 准备盖章页
```bash
python -m GaiZhangYe.core.entrypoints.cli_start --prepare
```

#### 盖章页覆盖
```bash
python -m GaiZhangYe.core.entrypoints.cli_start --cover
```

#### 批量Word转PDF
```bash
python -m GaiZhangYe.core.entrypoints.cli_start --convert
```

### 2. Web界面

```bash
python -m GaiZhangYe.web.app
```

然后在浏览器中访问 http://localhost:5000

## 项目结构

```
GaiZhangYe/
├── business_data/       # 业务数据
├── core/               # 核心功能
│   └── entrypoints/    # 入口文件
├── utils/              # 工具函数
└── web/                # Web界面
```

## 配置

配置文件采用.env格式，参考.env.example创建.env文件：

```bash
cp .env.example .env
```

## 开发

### 安装开发依赖

```bash
pip install -e .[dev]
```

### 运行测试

```bash
pytest
```

### 代码规范检查

```bash
ruff check .
```

## 构建

```bash
pyinstaller GaiZhangYe.spec
```

构建完成后，可执行文件将生成在dist/目录下。

## 许可证

MIT License

## 作者

Your Name - your@email.com

## 更新日志

### v0.1.0 (2025-12-19)
- 初始版本
- 实现核心功能：准备盖章页、盖章页覆盖、批量Word转PDF
- 支持CLI和Web两种操作方式