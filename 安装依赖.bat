@echo off
chcp 65001 >nul
echo ==========================================
echo 安装行政处罚文档生成工具依赖
echo ==========================================
echo.

echo 正在安装必要的 Python 库...
echo.

pip install docxtpl pandas openpyxl python-docx

echo.
if errorlevel 1 (
    echo 安装失败，请检查网络连接或手动运行：
    echo pip install docxtpl pandas openpyxl python-docx python-docx
) else (
    echo ==========================================
    echo 安装成功！
    echo ==========================================
    echo.
    echo 现在可以运行"开始生成.bat"了
)

pause
