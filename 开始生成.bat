@echo off
chcp 65001 >nul
echo ==========================================
echo 行政处罚文档生成工具
echo ==========================================
echo.

REM 检查是否存在打包后的 exe
if exist "行政处罚生成工具.exe" (
    echo 正在启动程序...
    行政处罚生成工具.exe
) else (
    echo 正在通过 Python 运行...
    python generate.py
)

pause
