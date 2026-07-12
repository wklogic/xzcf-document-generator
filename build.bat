@echo off
title Building render.exe
echo ==========================================
echo   Template Renderer - Build Script
echo ==========================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Install Python 3.8+
    pause
    exit /b 1
)

echo [1/4] Installing dependencies...
pip install docxtpl pandas openpyxl python-docx pyinstaller -q
if errorlevel 1 (
    echo ERROR: Dependency install failed
    pause
    exit /b 1
)
echo Done.
echo.

echo [2/4] Building render.exe...
pyinstaller -F --clean render.py --name render --distpath . -y ^
  --exclude-module torch ^
  --exclude-module torchvision ^
  --exclude-module scipy ^
  --exclude-module sqlalchemy ^
  --exclude-module onnxruntime ^
  --exclude-module tensorflow ^
  --exclude-module PIL ^
  --exclude-module cryptography ^
  --exclude-module pygments ^
  --exclude-module cffi
if errorlevel 1 (
    echo ERROR: Build failed
    pause
    exit /b 1
)
echo Done.
echo.

echo [3/4] Cleaning up...
if exist "build" rd /s /q "build"
if exist "render.spec" del /f /q "render.spec"
echo Done.
echo.

echo [4/4] Creating launcher...
echo @echo off > "开始生成.bat"
echo title Template Renderer >> "开始生成.bat"
echo if exist "render.exe" (render.exe) else (python render.py) >> "开始生成.bat"
echo pause >> "开始生成.bat"
echo Done.
echo.

echo ==========================================
echo   Build Complete!
echo ==========================================
echo.
echo   render.exe is ready.
echo   Copy this folder to distribute:
echo     - render.exe
echo     - config.json
echo     - templates\
echo     - data\
echo     - output\
echo     - 开始生成.bat
echo.
pause
