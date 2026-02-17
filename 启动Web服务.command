#!/bin/bash

# 获取脚本所在目录
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$DIR"

echo "================================"
echo "  公文格式调整工具 - Web版"
echo "================================"
echo ""

# 检查Python
if ! command -v python3 &> /dev/null; then
    echo "❌ 未找到 Python3，请先安装 Python"
    exit 1
fi

# 检查依赖
echo "🔍 检查依赖..."
if ! python3 -c "import flask" &> /dev/null; then
    echo "📦 安装依赖..."
    pip3 install -r requirements_web.txt
fi

echo ""
echo "🚀 启动Web服务..."
echo ""

# 启动Flask应用
python3 app.py
