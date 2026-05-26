@echo off
chcp 65001 >nul
echo ==========================================
echo 通用模板渲染工具 - 打包脚本
echo ==========================================
echo.

REM 检查 Python 是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误：未检测到 Python，请先安装 Python 3.8 或更高版本
    pause
    exit /b 1
)

echo [1/4] 正在安装依赖...
pip install docxtpl pandas openpyxl pyinstaller -q
if errorlevel 1 (
    echo 错误：依赖安装失败
    pause
    exit /b 1
)
echo 依赖安装完成！
echo.

echo [2/4] 正在打包程序...
pyinstaller -F --clean render.py --name render --distpath . -y
if errorlevel 1 (
    echo 错误：打包失败
    pause
    exit /b 1
)
echo 打包完成！
echo.

echo [3/4] 正在清理临时文件...
if exist "build" rd /s /q "build"
if exist "render.spec" del /f /q "render.spec"
echo 清理完成！
echo.

echo [4/4] 创建启动脚本...
echo @echo off > "开始生成.bat"
echo chcp 65001 ^>nul >> "开始生成.bat"
echo echo 正在启动通用模板渲染工具... >> "开始生成.bat"
echo render.exe >> "开始生成.bat"
echo pause >> "开始生成.bat"
echo 启动脚本创建完成！
echo.

echo ==========================================
echo 打包成功！
echo ==========================================
echo.
echo 生成的文件：
echo   - render.exe （主程序）
echo   - 开始生成.bat （双击运行此文件）
echo.
echo 使用步骤：
echo   1. 将模板文件放入"templates"文件夹
echo   2. 将数据文件放入"data"文件夹
echo   3. 双击"开始生成.bat"运行程序
echo   4. 生成的文档在"output"文件夹中
echo.
pause
