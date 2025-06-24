#!/usr/bin/env python3
"""
Docling 安装和功能测试脚本
用于验证 docling 库是否正确安装并能正常工作
"""

import sys
from pathlib import Path

def test_import():
    """测试 docling 库导入"""
    print("正在测试 docling 库导入...")
    try:
        from docling.document_converter import DocumentConverter
        from docling.datamodel.base_models import InputFormat
        from docling.datamodel.pipeline_options import PdfPipelineOptions
        from docling.document_converter import PdfFormatOption
        print("✅ Docling 库导入成功！")
        return True
    except ImportError as e:
        print(f"❌ Docling 库导入失败: {e}")
        print("请运行: uv pip install docling")
        return False

def test_converter_creation():
    """测试创建文档转换器"""
    print("\n正在测试创建文档转换器...")
    try:
        from docling.document_converter import DocumentConverter
        from docling.datamodel.base_models import InputFormat
        from docling.datamodel.pipeline_options import PdfPipelineOptions
        from docling.document_converter import PdfFormatOption
        
        # 创建基本转换器
        converter = DocumentConverter()
        print("✅ 基本转换器创建成功！")
        
        # 创建带配置的转换器
        pipeline_options = PdfPipelineOptions()
        pipeline_options.do_ocr = True
        pipeline_options.do_table_structure = True
        
        advanced_converter = DocumentConverter(
            format_options={
                InputFormat.PDF: PdfFormatOption(
                    pipeline_options=pipeline_options
                )
            }
        )
        print("✅ 高级转换器创建成功！")
        return True
        
    except Exception as e:
        print(f"❌ 转换器创建失败: {e}")
        return False

def test_online_pdf_conversion():
    """测试在线 PDF 转换（使用 docling 官方示例）"""
    print("\n正在测试在线 PDF 转换...")
    try:
        from docling.document_converter import DocumentConverter
        
        # 使用 docling 官方文档中的示例 PDF
        source = "https://arxiv.org/pdf/2408.09869"
        converter = DocumentConverter()
        
        print(f"正在下载并转换: {source}")
        print("这可能需要几分钟时间，特别是首次运行时...")
        
        result = converter.convert(source)
        markdown_content = result.document.export_to_markdown()
        
        if markdown_content and len(markdown_content) > 100:
            print("✅ 在线 PDF 转换成功！")
            print(f"转换内容长度: {len(markdown_content)} 字符")
            
            # 显示前200个字符作为预览
            preview = markdown_content[:200].replace('\n', ' ')
            print(f"内容预览: {preview}...")
            
            # 可选：保存到文件
            output_file = Path("test_output.md")
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(markdown_content)
            print(f"✅ 转换结果已保存到: {output_file}")
            
            return True
        else:
            print("❌ 转换结果为空或过短")
            return False
            
    except Exception as e:
        print(f"❌ 在线 PDF 转换失败: {e}")
        print("这可能是网络问题或首次下载模型时间较长")
        return False

def main():
    """主测试函数"""
    print("=== Docling 安装和功能测试 ===\n")
    
    # 测试导入
    if not test_import():
        sys.exit(1)
    
    # 测试转换器创建
    if not test_converter_creation():
        sys.exit(1)
    
    # 询问是否进行在线测试
    print("\n是否进行在线 PDF 转换测试？")
    print("注意：这需要网络连接，首次运行时会下载模型文件（可能需要几分钟）")
    choice = input("输入 'y' 继续，其他键跳过: ").lower().strip()
    
    if choice == 'y':
        test_online_pdf_conversion()
    else:
        print("跳过在线测试")
    
    print("\n=== 测试完成 ===")
    print("如果所有测试都通过，您可以开始使用 pdf_to_markdown_converter.py 脚本了！")

if __name__ == '__main__':
    main() 