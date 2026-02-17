@echo off
chcp 65001 >nul
cd /d %~dp0

echo ================================
echo   公文格式调整工具 - Web版
echo ================================
echo.

REM 检查Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 未找到 Python，请先安装 Python
    pause
    exit /b 1
)

REM 检查依赖
echo 🔍 检查依赖...
python -c "import flask" >nul 2>&1
if errorlevel 1 (
    echo 📦 安装依赖...
    pip install -r requirements_web.txt
)

echo.
echo 🚀 启动Web服务...
echo.

REM 启动Flask应用
python app.py

pause
