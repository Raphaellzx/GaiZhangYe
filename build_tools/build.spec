# build_tools/build.spec
import sys
import os
from pathlib import Path
from PyInstaller.utils.hooks import collect_all

# ========== 核心路径配置（适配项目目录） ==========
ROOT_DIR = Path(__file__).resolve().parent.parent  # 项目根目录
SRC_DIR = ROOT_DIR / "GaiZhangYe"                  # 源码目录
ENTRY_FILE = SRC_DIR / "entrypoints" / "gui.py"    # 启动入口
DIST_TEMP_DIR = ROOT_DIR / "dist_temp"             # 打包临时输出目录
BUILD_TEMP_DIR = ROOT_DIR / "build"                # 编译临时目录

# 创建临时目录（避免打包时目录不存在）
DIST_TEMP_DIR.mkdir(exist_ok=True)
BUILD_TEMP_DIR.mkdir(exist_ok=True)

# ========== 依赖收集（解决pywin32/pymupdf等问题） ==========
datas = []
binaries = []
hiddenimports = [
    "win32com", "win32com.client", "pythoncom",
    "fitz", "PIL", "tkinter", "tkinter.ttk",
    "pydantic_settings", "dotenv"
]

# 收集pymupdf/pillow依赖
for pkg in ["fitz", "PIL"]:
    pkg_datas, pkg_binaries, pkg_hidden = collect_all(pkg)
    datas += pkg_datas
    binaries += pkg_binaries
    hiddenimports += pkg_hidden

# ========== PyInstaller核心配置 ==========
a = Analysis(
    [str(ENTRY_FILE)],
    pathex=[str(ROOT_DIR)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
)

pyz = PYZ(a.pure, a.zipped_data)

# ========== EXE输出配置（适配调试/发布） ==========
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    name="GaiZhangYe",
    debug=False,                  # 调试时改为True
    console=False,                # 调试时改为True（显示控制台），发布时False
    strip=False,
    upx=True,                     # 压缩体积（需安装UPX）
    runtime_tmpdir=None,
    disable_windowed_traceback=False,
    target_arch=None,
    icon=None,                    # 可选：添加图标，如 ROOT_DIR / "build_tools" / "icon.ico"
    # 输出到临时目录
    distpath=str(DIST_TEMP_DIR),
    workpath=str(BUILD_TEMP_DIR),
)