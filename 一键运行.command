#!/bin/bash
# 公文格式调整工具 - macOS 一键运行脚本

# 获取脚本所在目录
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

# 设置终端标题
echo -e "\033]0;公文格式调整工具\007"

# 清屏
clear

echo "=================================================="
echo "  📄 公文格式调整工具 - macOS 版"
echo "=================================================="
echo ""

# 检查 Python 是否安装
if ! command -v python3 &> /dev/null
then
    echo "❌ 错误：未检测到 Python 3"
    echo "请先安装 Python 3："
    echo "  https://www.python.org/downloads/"
    echo ""
    read -p "按回车键退出..."
    exit 1
fi

echo "✅ Python 版本: $(python3 --version)"
echo ""

# 检查依赖是否安装
if ! python3 -c "import docx" 2>/dev/null; then
    echo "⚠️  未检测到 python-docx 库"
    echo "正在自动安装..."
    echo ""
    pip3 install python-docx
    echo ""
fi

echo "=================================================="
echo ""

# 运行主程序
python3 gongwen_formatter_cli.py

# 结束时暂停（防止窗口关闭）
echo ""
read -p "按回车键退出..."
