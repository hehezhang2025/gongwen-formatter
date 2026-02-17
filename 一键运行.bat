@echo off
chcp 65001 >nul
REM 公文格式调整工具 - Windows 一键运行脚本

title 公文格式调整工具
color 0A
cls

echo ==================================================
echo   📄 公文格式调整工具 - Windows 版
echo ==================================================
echo.

REM 检查 Python 是否安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ 错误：未检测到 Python 3
    echo 请先安装 Python 3：
    echo   https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo ✅ Python 版本:
python --version
echo.

REM 检查依赖是否安装
python -c "import docx" >nul 2>&1
if %errorlevel% neq 0 (
    echo ⚠️  未检测到 python-docx 库
    echo 正在自动安装...
    echo.
    pip install python-docx
    echo.
)

echo ==================================================
echo.

REM 运行主程序
python gongwen_formatter_cli.py

REM 结束时暂停
echo.
pause
