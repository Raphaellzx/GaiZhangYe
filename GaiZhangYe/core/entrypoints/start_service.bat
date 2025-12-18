@echo off
title 盖章页工具 - 服务启动

echo ====================================
echo 盖章页工具 HTML 可视化服务启动脚本
echo ====================================
echo.

:: 1. 杀死所有运行中的Python进程
echo [1/3] 正在杀死旧的Python进程...
taskkill /F /IM python.exe >nul 2>&1

:: 2. 清除浏览器缓存（仅Edge）
echo [2/3] 正在清除浏览器缓存...
:: 清除Edge浏览器缓存的命令（需要管理员权限，可选）
:: RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 255

:: 3. 启动服务
echo [3/3] 正在启动HTML服务...
echo.
echo 服务将在 http://localhost:5001 启动
echo 浏览器将在2秒后自动打开
echo 按 Ctrl+C 停止服务
echo.

:: 使用Python启动服务
python -m GaiZhangYe.entrypoints.start_web

echo.
echo 服务已停止
pause
