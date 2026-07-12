@echo off
title Template Renderer
echo ==========================================
echo   Template Renderer v1.0
echo ==========================================
echo.

if exist "render.exe" (
    echo Starting render.exe ...
    render.exe
) else (
    echo render.exe not found, trying Python ...
    python render.py
)

pause
