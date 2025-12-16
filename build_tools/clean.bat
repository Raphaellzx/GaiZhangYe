@echo off
chcp 65001 >nul 2>&1
setlocal

set "PROJECT_ROOT=%~dp0.."

:: 清理PyInstaller临时文件
echo 清理打包临时文件...
if exist "%PROJECT_ROOT%\dist_temp" rd /s /q "%PROJECT_ROOT%\dist_temp"
if exist "%PROJECT_ROOT%\build" rd /s /q "%PROJECT_ROOT%\build"
if exist "%PROJECT_ROOT%\GaiZhangYe.spec" del /f /q "%PROJECT_ROOT%\GaiZhangYe.spec"

:: 可选：清理日志和业务数据（调试时用）
:: if exist "%PROJECT_ROOT%\logs" rd /s /q "%PROJECT_ROOT%\logs"
:: if exist "%PROJECT_ROOT%\business_data" rd /s /q "%PROJECT_ROOT%\business_data"

echo 清理完成！
pause