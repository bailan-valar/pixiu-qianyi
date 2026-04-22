#!/bin/bash
# 貔貅记账转换工具启动脚本

echo "正在启动貔貅记账转换工具..."

# 检查Python是否安装
if ! command -v python3 &> /dev/null; then
    echo "错误: 未找到Python3，请先安装Python3"
    exit 1
fi

# 检查必要的依赖
python3 -c "import tkinter" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "错误: tkinter模块未安装"
    echo "在macOS上，通常已经包含tkinter"
    echo "在Linux上，请运行: sudo apt-get install python3-tk"
    exit 1
fi

python3 -c "import pandas" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "正在安装pandas..."
    pip3 install pandas
fi

# 启动应用
python3 pixiu_converter.py
