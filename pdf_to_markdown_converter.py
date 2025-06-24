#!/usr/bin/env python3
"""
PDF批量转Markdown转换器
使用Docling库将文件夹中的所有PDF文件转换为Markdown格式
"""

import os
import sys
import argparse
from pathlib import Path
from typing import List, Optional
import logging
from datetime import datetime

# 导入docling相关模块
try:
    from docling.document_converter import DocumentConverter
    from docling.datamodel.base_models import InputFormat
    from docling.datamodel.pipeline_options import PdfPipelineOptions
    from docling.document_converter import PdfFormatOption
except ImportError as e:
    print(f"错误：无法导入docling库。请确保已安装docling: {e}")
    print("运行命令：uv pip install docling")
    sys.exit(1)

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('pdf_to_markdown_conversion.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class PdfToMarkdownConverter:
    """PDF转Markdown转换器类"""
    
    def __init__(self, enable_ocr: bool = True, enable_table_structure: bool = True):
        """
        初始化转换器
        
        Args:
            enable_ocr: 是否启用OCR识别扫描版PDF
            enable_table_structure: 是否启用表格结构识别
        """
        self.enable_ocr = enable_ocr
        self.enable_table_structure = enable_table_structure
        self.converter = self._create_converter()
        
    def _create_converter(self) -> DocumentConverter:
        """创建文档转换器"""
        try:
            # 配置PDF处理选项
            pipeline_options = PdfPipelineOptions()
            pipeline_options.do_ocr = self.enable_ocr
            pipeline_options.do_table_structure = self.enable_table_structure
            
            if self.enable_table_structure:
                # 启用表格单元格匹配以获得更好的表格结构
                pipeline_options.table_structure_options.do_cell_matching = True
            
            # 创建转换器
            converter = DocumentConverter(
                format_options={
                    InputFormat.PDF: PdfFormatOption(
                        pipeline_options=pipeline_options
                    )
                }
            )
            
            logger.info("文档转换器初始化成功")
            return converter
            
        except Exception as e:
            logger.error(f"创建转换器时出错: {str(e)}")
            raise
    
    def convert_pdf_to_markdown(self, pdf_path: Path, output_dir: Path, 
                              overwrite: bool = False) -> bool:
        """
        将单个PDF文件转换为Markdown
        
        Args:
            pdf_path: PDF文件路径
            output_dir: 输出目录
            overwrite: 是否覆盖已存在的文件
            
        Returns:
            转换是否成功
        """
        try:
            # 检查PDF文件是否存在
            if not pdf_path.exists():
                logger.error(f"PDF文件不存在: {pdf_path}")
                return False
            
            # 创建输出目录
            output_dir.mkdir(parents=True, exist_ok=True)
            
            # 生成输出文件路径
            markdown_file = output_dir / f"{pdf_path.stem}.md"
            
            # 检查输出文件是否已存在
            if markdown_file.exists() and not overwrite:
                logger.warning(f"Markdown文件已存在，跳过: {markdown_file}")
                return True
            
            logger.info(f"开始转换: {pdf_path.name} -> {markdown_file.name}")
            
            # 执行转换
            result = self.converter.convert(str(pdf_path))
            
            # 导出为Markdown
            markdown_content = result.document.export_to_markdown()
            
            # 保存Markdown文件
            with open(markdown_file, 'w', encoding='utf-8') as f:
                f.write(markdown_content)
            
            # 检查生成的文件
            if markdown_file.exists() and markdown_file.stat().st_size > 0:
                logger.info(f"转换成功: {pdf_path.name} ({markdown_file.stat().st_size} bytes)")
                return True
            else:
                logger.error(f"转换失败: 生成的Markdown文件为空或不存在")
                return False
                
        except Exception as e:
            logger.error(f"转换PDF文件时出错 {pdf_path}: {str(e)}")
            return False

def find_pdf_files(directory: str, recursive: bool = True) -> List[Path]:
    """查找指定目录下的所有PDF文件"""
    directory_path = Path(directory)
    
    if not directory_path.exists():
        logger.error(f"目录不存在: {directory}")
        return []
    
    if not directory_path.is_dir():
        logger.error(f"路径不是目录: {directory}")
        return []
    
    if recursive:
        pdf_files = list(directory_path.rglob("*.pdf"))
    else:
        pdf_files = list(directory_path.glob("*.pdf"))
    
    # 过滤掉隐藏文件和临时文件
    pdf_files = [f for f in pdf_files if not f.name.startswith('.') and not f.name.startswith('~')]
    
    pdf_files = sorted(pdf_files)
    logger.info(f"在目录 {directory} 中找到 {len(pdf_files)} 个PDF文件")
    
    return pdf_files

def batch_convert(input_dir: str, output_dir: Optional[str] = None, 
                 recursive: bool = True, overwrite: bool = False,
                 enable_ocr: bool = True, enable_table_structure: bool = True) -> dict:
    """批量转换PDF文件为Markdown"""
    
    # 查找PDF文件
    pdf_files = find_pdf_files(input_dir, recursive)
    if not pdf_files:
        logger.warning("没有找到PDF文件")
        return {'success': 0, 'failed': 0, 'skipped': 0}
    
    # 确定输出目录
    if output_dir:
        output_path = Path(output_dir)
    else:
        output_path = Path(input_dir) / "markdown_output"
    
    # 创建转换器
    try:
        converter = PdfToMarkdownConverter(
            enable_ocr=enable_ocr,
            enable_table_structure=enable_table_structure
        )
    except Exception as e:
        logger.error(f"创建转换器失败: {str(e)}")
        return {'success': 0, 'failed': 0, 'skipped': 0}
    
    logger.info(f"开始批量转换，输出目录: {output_path}")
    logger.info(f"OCR启用: {enable_ocr}, 表格结构识别启用: {enable_table_structure}")
    
    results = {'success': 0, 'failed': 0, 'skipped': 0}
    
    for i, pdf_file in enumerate(pdf_files, 1):
        logger.info(f"处理进度: {i}/{len(pdf_files)}")
        
        try:
            # 保持相对路径结构
            relative_path = pdf_file.relative_to(Path(input_dir))
            output_subdir = output_path / relative_path.parent
            
            # 转换文档
            if converter.convert_pdf_to_markdown(pdf_file, output_subdir, overwrite):
                results['success'] += 1
            else:
                results['failed'] += 1
                
        except Exception as e:
            logger.error(f"处理文件时出错 {pdf_file}: {str(e)}")
            results['failed'] += 1
    
    return results

def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description='批量将PDF文档转换为Markdown (使用Docling)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  python pdf_to_markdown_converter.py /path/to/pdf/files
  python pdf_to_markdown_converter.py /path/to/pdf/files -o /path/to/output
  python pdf_to_markdown_converter.py /path/to/pdf/files --no-recursive --no-ocr
        """
    )
    
    parser.add_argument(
        'input_dir',
        help='包含PDF文档的输入目录'
    )
    
    parser.add_argument(
        '-o', '--output-dir',
        help='Markdown输出目录（默认为输入目录下的markdown_output文件夹）',
        default=None
    )
    
    parser.add_argument(
        '--no-recursive',
        action='store_true',
        help='不递归搜索子目录中的PDF文件'
    )
    
    parser.add_argument(
        '--overwrite',
        action='store_true',
        help='覆盖已存在的Markdown文件'
    )
    
    parser.add_argument(
        '--no-ocr',
        action='store_true',
        help='禁用OCR功能（扫描版PDF可能转换效果不佳）'
    )
    
    parser.add_argument(
        '--no-table-structure',
        action='store_true',
        help='禁用表格结构识别'
    )
    
    args = parser.parse_args()
    
    # 验证输入目录
    if not os.path.exists(args.input_dir):
        logger.error(f"输入目录不存在: {args.input_dir}")
        sys.exit(1)
    
    if not os.path.isdir(args.input_dir):
        logger.error(f"输入路径不是目录: {args.input_dir}")
        sys.exit(1)
    
    logger.info("="*60)
    logger.info("PDF批量转Markdown转换器 (基于Docling)")
    logger.info("="*60)
    logger.info(f"输入目录: {args.input_dir}")
    logger.info(f"输出目录: {args.output_dir or '输入目录/markdown_output'}")
    logger.info(f"递归搜索: {not args.no_recursive}")
    logger.info(f"覆盖现有文件: {args.overwrite}")
    logger.info(f"OCR功能: {not args.no_ocr}")
    logger.info(f"表格结构识别: {not args.no_table_structure}")
    logger.info("="*60)
    
    # 记录开始时间
    start_time = datetime.now()
    
    # 执行批量转换
    results = batch_convert(
        input_dir=args.input_dir,
        output_dir=args.output_dir,
        recursive=not args.no_recursive,
        overwrite=args.overwrite,
        enable_ocr=not args.no_ocr,
        enable_table_structure=not args.no_table_structure
    )
    
    # 记录结束时间
    end_time = datetime.now()
    duration = end_time - start_time
    
    # 输出结果统计
    logger.info("="*60)
    logger.info("转换完成!")
    logger.info(f"耗时: {duration}")
    logger.info(f"成功转换: {results['success']} 个文件")
    logger.info(f"转换失败: {results['failed']} 个文件")
    logger.info(f"跳过文件: {results['skipped']} 个文件")
    logger.info("="*60)
    
    if results['failed'] > 0:
        logger.warning("部分文件转换失败，请查看日志了解详情")
        sys.exit(1)

if __name__ == '__main__':
    main() 