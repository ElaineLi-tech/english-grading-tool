#!/bin/bash
# 英语作业自动批改工具 - 启动脚本

cd "$(dirname "$0")"

echo "=========================================="
echo "  英语作业自动批改工具"
echo "=========================================="
echo ""

# 检查Python
if ! command -v python3 &> /dev/null; then
    echo "❌ 未找到 Python3，请先安装 Python"
    exit 1
fi

# 检查依赖
echo "📦 检查依赖..."
python3 -c "import flask" 2>/dev/null || {
    echo "  安装 Flask..."
    pip3 install flask --break-system-packages
}

python3 -c "import docx" 2>/dev/null || {
    echo "  安装 python-docx..."
    pip3 install python-docx --break-system-packages
}

python3 -c "from PIL import Image" 2>/dev/null || {
    echo "  安装 Pillow..."
    pip3 install Pillow --break-system-packages
}

python3 -c "import requests" 2>/dev/null || {
    echo "  安装 requests..."
    pip3 install requests --break-system-packages
}

echo ""
echo "🚀 启动服务..."
echo ""
echo "  🌐 访问地址：http://localhost:5005"
echo "  按 Ctrl+C 停止服务"
echo ""

python3 app.py
