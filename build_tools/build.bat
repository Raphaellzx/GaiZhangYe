@echo off
chcp 65001 >nul 2>&1  # 解决中文乱码
setlocal enabledelayedexpansion

:: ========== 配置项（可根据需要修改） ==========
set "PROJECT_ROOT=%~dp0.."  # 项目根目录（build_tools的上级）
set "SPEC_FILE=%PROJECT_ROOT%\build_tools\build.spec"
set "DIST_TEMP=%PROJECT_ROOT%\dist_temp"
set "DIST_RELEASE=%PROJECT_ROOT%\dist_release"
set "VERSION=v0.1.0"        # 版本号（打包时自动带入）
set "RELEASE_DIR=%DIST_RELEASE%\GaiZhangYe_%VERSION%"

:: ========== 步骤1：清理旧打包文件 ==========
echo [1/5] 清理旧打包文件...
if exist "%DIST_TEMP%" rd /s /q "%DIST_TEMP%"
if exist "%PROJECT_ROOT%\build" rd /s /q "%PROJECT_ROOT%\build"
if exist "%RELEASE_DIR%" rd /s /q "%RELEASE_DIR%"
mkdir "%RELEASE_DIR%" >nul 2>&1

:: ========== 步骤2：激活uv虚拟环境并安装依赖（确保依赖最新） ==========
echo [2/5] 检查依赖并激活环境...
cd /d "%PROJECT_ROOT%"
uv install --dev >nul 2>&1 || (
    echo 依赖安装失败！
    pause
    exit /b 1
)

:: ========== 步骤3：执行PyInstaller打包 ==========
echo [3/5] 开始打包exe...
uv run pyinstaller "%SPEC_FILE%" || (
    echo 打包失败！
    pause
    exit /b 1
)

:: ========== 步骤4：整理发布目录（复制exe+配置文件） ==========
echo [4/5] 整理发布文件...
copy "%DIST_TEMP%\GaiZhangYe.exe" "%RELEASE_DIR%\" >nul 2>&1
copy "%PROJECT_ROOT%\.env.example" "%RELEASE_DIR%\.env.example" >nul 2>&1
copy "%PROJECT_ROOT%\README.md" "%RELEASE_DIR%\" >nul 2>&1
:: 复制.env（若存在，用户可自定义配置）
if exist "%PROJECT_ROOT%\.env" copy "%PROJECT_ROOT%\.env" "%RELEASE_DIR%\" >nul 2>&1

:: ========== 步骤5：完成提示 ==========
echo [5/5] 打包完成！
echo 发布包路径：%RELEASE_DIR%
echo 双击 GaiZhangYe.exe 即可运行
pause