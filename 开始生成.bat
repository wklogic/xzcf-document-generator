@echo off
chcp 65001 >nul
echo ==========================================
echo 通用模板渲染工具
echo ==========================================
echo.

REM 检查是否存在打包后的 exe
if exist "render.exe" (
    echo 正在启动程序...
    render.exe
) else (
    echo 正在通过 Python 运行...
    python render.py
)

pause
