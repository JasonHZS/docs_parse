#!/bin/bash

# 综合文档转Markdown转换器运行脚本
# 自动激活docling虚拟环境并运行转换

echo "正在启动综合文档转Markdown转换器..."
echo "目标目录: /Users/oasis/Documents/work/运营知识库"
echo "=========================================="

# 检查虚拟环境是否存在
if [ ! -d "venv_docling" ]; then
    echo "错误: 找不到docling虚拟环境目录 venv_docling"
    echo "请先创建虚拟环境: python -m venv venv_docling"
    exit 1
fi

# 激活虚拟环境
echo "正在激活docling虚拟环境..."
source venv_docling/bin/activate

# 检查是否成功激活
if [ "$VIRTUAL_ENV" == "" ]; then
    echo "错误: 无法激活虚拟环境"
    exit 1
fi

echo "虚拟环境已激活: $VIRTUAL_ENV"

# 检查docling是否已安装
echo "检查docling库..."
python -c "import docling" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "docling库未安装，正在安装..."
    pip install docling
    if [ $? -ne 0 ]; then
        echo "错误: docling安装失败"
        exit 1
    fi
fi

# 检查LibreOffice是否可用
echo "检查LibreOffice..."
if ! command -v libreoffice &> /dev/null && ! command -v soffice &> /dev/null; then
    echo "警告: 未检测到LibreOffice，Word文件转换可能失败"
    echo "请安装LibreOffice: brew install --cask libreoffice"
    echo "是否继续？(y/n)"
    read -r response
    if [[ "$response" != "y" && "$response" != "Y" ]]; then
        echo "取消转换"
        exit 1
    fi
fi

# 运行转换脚本
echo "开始转换..."
python comprehensive_to_markdown_converter.py "/Users/oasis/Documents/work/运营知识库" "$@"

# 检查转换结果
if [ $? -eq 0 ]; then
    echo "=========================================="
    echo "转换完成! 结果保存在: /Users/oasis/Documents/work/运营知识库/markdown_output"
    echo "日志文件: comprehensive_conversion.log"
else
    echo "转换过程中出现错误，请查看日志文件了解详情"
fi

# 保持虚拟环境激活状态（可选）
echo "虚拟环境仍处于激活状态，您可以继续使用docling相关功能"
echo "要退出虚拟环境，请输入: deactivate" 