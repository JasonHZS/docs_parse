#!/usr/bin/env python3
"""
综合文档转Markdown转换器
将指定目录下的Word和Excel文件转换为Markdown格式
- Word文件(.doc/.docx)：先转PDF，再转Markdown
- Excel文件(.xlsx/.xls)：每个sheet转换为单独的Markdown文件
"""

import os
import sys
import argparse
import subprocess
import shutil
from pathlib import Path
from typing import List, Optional, Dict
import logging
from datetime import datetime
import re

# 导入docling相关模块
try:
    from docling.document_converter import DocumentConverter
    from docling.datamodel.base_models import InputFormat
    from docling.datamodel.pipeline_options import PdfPipelineOptions, EasyOcrOptions
    from docling.document_converter import PdfFormatOption
except ImportError as e:
    print(f"错误：无法导入docling库。请确保已激活docling虚拟环境并安装了docling: {e}")
    print("运行命令：source venv_docling/bin/activate && pip install docling")
    sys.exit(1)

# 导入Excel处理相关模块
try:
    import pandas as pd
    import openpyxl
except ImportError as e:
    print(f"错误：无法导入Excel处理库。请安装pandas和openpyxl: {e}")
    print("运行命令：pip install pandas openpyxl")
    sys.exit(1)

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('comprehensive_conversion.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class ComprehensiveToMarkdownConverter:
    """综合文档转Markdown转换器"""
    
    def __init__(self, enable_ocr: bool = True, enable_table_structure: bool = True):
        """
        初始化转换器
        
        Args:
            enable_ocr: 是否启用OCR识别
            enable_table_structure: 是否启用表格结构识别
        """
        self.enable_ocr = enable_ocr
        self.enable_table_structure = enable_table_structure
        self.converter = self._create_converter()
        self.temp_pdf_dir = Path("temp_pdf_files")
        self.temp_pdf_dir.mkdir(exist_ok=True)
        
        # 检查LibreOffice是否可用
        self.libreoffice_available = self._check_libreoffice()
        
    def _create_converter(self) -> DocumentConverter:
        """创建docling文档转换器"""
        try:
            # 配置处理选项
            pipeline_options = PdfPipelineOptions()
            pipeline_options.do_ocr = self.enable_ocr
            pipeline_options.do_table_structure = self.enable_table_structure
            
            if self.enable_table_structure:
                pipeline_options.table_structure_options.do_cell_matching = True
            
            # 创建转换器，支持PDF和Excel格式
            converter = DocumentConverter(
                format_options={
                    InputFormat.PDF: PdfFormatOption(
                        pipeline_options=pipeline_options
                    )
                }
            )
            
            logger.info("Docling转换器初始化成功")
            return converter
            
        except Exception as e:
            logger.error(f"创建docling转换器时出错: {str(e)}")
            raise
    
    def _check_libreoffice(self) -> bool:
        """检查LibreOffice是否可用"""
        if shutil.which('libreoffice') or shutil.which('soffice'):
            logger.info("检测到LibreOffice")
            return True
        else:
            logger.warning("未检测到LibreOffice，Word文件转换可能失败")
            return False
    
    def _convert_word_to_pdf(self, word_file: Path, temp_dir: Path) -> Optional[Path]:
        """使用LibreOffice将Word文件转换为PDF"""
        if not self.libreoffice_available:
            logger.error("LibreOffice不可用，无法转换Word文件")
            return None
        
        try:
            # 使用LibreOffice的无头模式转换
            cmd_name = 'soffice' if shutil.which('soffice') else 'libreoffice'
            
            cmd = [
                cmd_name,
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', str(temp_dir),
                str(word_file)
            ]
            
            logger.info(f"执行Word转PDF: {word_file.name}")
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=120  # 2分钟超时
            )
            
            if result.returncode == 0:
                pdf_file = temp_dir / f"{word_file.stem}.pdf"
                if pdf_file.exists() and pdf_file.stat().st_size > 0:
                    logger.info(f"Word转PDF成功: {word_file.name} -> {pdf_file.name}")
                    return pdf_file
            
            logger.error(f"LibreOffice转换失败: {result.stderr}")
            return None
            
        except subprocess.TimeoutExpired:
            logger.error(f"LibreOffice转换超时: {word_file.name}")
            return None
        except Exception as e:
            logger.error(f"Word转PDF出错: {str(e)}")
            return None
    
    def _convert_to_markdown_with_docling(self, file_path: Path) -> Optional[str]:
        """使用docling将文件转换为Markdown"""
        try:
            logger.info(f"使用docling转换为Markdown: {file_path.name}")
            
            # 执行转换
            result = self.converter.convert(str(file_path))
            
            # 导出为Markdown
            markdown_content = result.document.export_to_markdown()
            
            if markdown_content:
                logger.info(f"Docling转换成功: {file_path.name}")
                return markdown_content
            else:
                logger.error(f"Docling转换结果为空: {file_path.name}")
                return None
                
        except Exception as e:
            logger.error(f"Docling转换出错 {file_path.name}: {str(e)}")
            return None
    
    def convert_word_to_markdown(self, word_file: Path, output_dir: Path) -> bool:
        """将Word文件转换为Markdown"""
        try:
            # 第一步：Word转PDF
            pdf_file = self._convert_word_to_pdf(word_file, self.temp_pdf_dir)
            if not pdf_file:
                return False
            
            # 第二步：PDF转Markdown
            markdown_content = self._convert_to_markdown_with_docling(pdf_file)
            if not markdown_content:
                return False
            
            # 第三步：保存Markdown文件
            output_dir.mkdir(parents=True, exist_ok=True)
            markdown_file = output_dir / f"{word_file.stem}.md"
            
            with open(markdown_file, 'w', encoding='utf-8') as f:
                f.write(markdown_content)
            
            logger.info(f"Word文件转换完成: {word_file.name} -> {markdown_file.name}")
            
            # 清理临时PDF文件
            try:
                pdf_file.unlink()
            except:
                pass
            
            return True
            
        except Exception as e:
            logger.error(f"Word转Markdown失败 {word_file.name}: {str(e)}")
            return False
    
    def _get_excel_sheet_names(self, excel_file: Path) -> List[str]:
        """获取Excel文件中的所有sheet名称"""
        try:
            workbook = openpyxl.load_workbook(excel_file, read_only=True)
            sheet_names = workbook.sheetnames
            workbook.close()
            logger.info(f"Excel文件 {excel_file.name} 包含 {len(sheet_names)} 个sheet: {sheet_names}")
            return sheet_names
        except Exception as e:
            logger.error(f"获取Excel sheet名称失败 {excel_file.name}: {str(e)}")
            return []
    
    def _convert_sheet_to_markdown(self, excel_file: Path, sheet_name: str) -> Optional[str]:
        """将Excel的单个sheet转换为Markdown"""
        try:
            # 读取sheet数据
            df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
            
            if df.empty:
                logger.warning(f"Sheet '{sheet_name}' 为空")
                return f"# {sheet_name}\n\n此sheet为空。\n"
            
            # 转换为Markdown表格
            markdown_content = f"# {sheet_name}\n\n"
            
            # 检查是否有数据
            if len(df) > 0 and len(df.columns) > 0:
                # 处理表头
                if len(df) > 0:
                    # 第一行作为表头
                    headers = df.iloc[0].fillna('').astype(str)
                    markdown_content += "| " + " | ".join(headers) + " |\n"
                    markdown_content += "| " + " | ".join(["---"] * len(headers)) + " |\n"
                    
                    # 数据行
                    for i in range(1, len(df)):
                        row = df.iloc[i].fillna('').astype(str)
                        markdown_content += "| " + " | ".join(row) + " |\n"
                else:
                    markdown_content += "| 数据 |\n| --- |\n"
            else:
                markdown_content += "此sheet不包含表格数据。\n"
            
            return markdown_content
            
        except Exception as e:
            logger.error(f"转换sheet '{sheet_name}' 失败: {str(e)}")
            return None
    
    def _sanitize_filename(self, filename: str) -> str:
        """清理文件名，移除或替换不安全的字符"""
        # 移除或替换不安全的字符
        sanitized = re.sub(r'[<>:"/\\|?*]', '_', filename)
        # 移除前后空格
        sanitized = sanitized.strip()
        # 如果文件名为空，使用默认名称
        if not sanitized:
            sanitized = "unnamed_sheet"
        return sanitized
    
    def convert_excel_to_markdown(self, excel_file: Path, output_dir: Path) -> bool:
        """将Excel文件的每个sheet转换为单独的Markdown文件"""
        try:
            # 获取所有sheet名称
            sheet_names = self._get_excel_sheet_names(excel_file)
            if not sheet_names:
                logger.error(f"无法获取Excel文件的sheet信息: {excel_file.name}")
                return False
            
            # 创建输出目录
            output_dir.mkdir(parents=True, exist_ok=True)
            
            success_count = 0
            total_sheets = len(sheet_names)
            
            # 为每个sheet创建单独的Markdown文件
            for sheet_name in sheet_names:
                try:
                    logger.info(f"处理sheet: {sheet_name}")
                    
                    # 转换sheet为Markdown
                    markdown_content = self._convert_sheet_to_markdown(excel_file, sheet_name)
                    if not markdown_content:
                        logger.warning(f"Sheet '{sheet_name}' 转换失败")
                        continue
                    
                    # 生成安全的文件名
                    safe_sheet_name = self._sanitize_filename(sheet_name)
                    markdown_filename = f"{excel_file.stem}-{safe_sheet_name}.md"
                    markdown_file = output_dir / markdown_filename
                    
                    # 保存Markdown文件
                    with open(markdown_file, 'w', encoding='utf-8') as f:
                        f.write(markdown_content)
                    
                    logger.info(f"Sheet转换完成: {sheet_name} -> {markdown_file.name}")
                    success_count += 1
                    
                except Exception as e:
                    logger.error(f"处理sheet '{sheet_name}' 时出错: {str(e)}")
                    continue
            
            if success_count > 0:
                logger.info(f"Excel文件转换完成: {excel_file.name} -> {success_count}/{total_sheets} 个sheet")
                return True
            else:
                logger.error(f"Excel文件转换失败: {excel_file.name} - 没有成功转换的sheet")
                return False
            
        except Exception as e:
            logger.error(f"Excel转Markdown失败 {excel_file.name}: {str(e)}")
            return False
    
    def cleanup(self):
        """清理临时文件"""
        try:
            if self.temp_pdf_dir.exists():
                shutil.rmtree(self.temp_pdf_dir)
                logger.info("清理临时文件完成")
        except Exception as e:
            logger.warning(f"清理临时文件失败: {str(e)}")

def find_target_files(directory: str, recursive: bool = True) -> Dict[str, List[Path]]:
    """查找目标文件（Word和Excel）"""
    directory_path = Path(directory)
    
    if not directory_path.exists():
        logger.error(f"目录不存在: {directory}")
        return {'word': [], 'excel': []}
    
    if not directory_path.is_dir():
        logger.error(f"路径不是目录: {directory}")
        return {'word': [], 'excel': []}
    
    word_extensions = ['*.doc', '*.docx']
    excel_extensions = ['*.xls', '*.xlsx']
    
    files = {'word': [], 'excel': []}
    
    # 查找Word文件
    for ext in word_extensions:
        if recursive:
            found_files = list(directory_path.rglob(ext))
        else:
            found_files = list(directory_path.glob(ext))
        files['word'].extend(found_files)
    
    # 查找Excel文件
    for ext in excel_extensions:
        if recursive:
            found_files = list(directory_path.rglob(ext))
        else:
            found_files = list(directory_path.glob(ext))
        files['excel'].extend(found_files)
    
    # 过滤隐藏文件和临时文件
    files['word'] = [f for f in files['word'] if not f.name.startswith('.') and not f.name.startswith('~')]
    files['excel'] = [f for f in files['excel'] if not f.name.startswith('.') and not f.name.startswith('~')]
    
    files['word'] = sorted(list(set(files['word'])))
    files['excel'] = sorted(list(set(files['excel'])))
    
    logger.info(f"找到Word文件: {len(files['word'])} 个")
    logger.info(f"找到Excel文件: {len(files['excel'])} 个")
    
    return files

def batch_convert_all(input_dir: str, output_dir: Optional[str] = None, 
                     recursive: bool = True, overwrite: bool = False,
                     enable_ocr: bool = True, enable_table_structure: bool = True) -> dict:
    """批量转换所有支持的文件为Markdown"""
    
    # 查找目标文件
    target_files = find_target_files(input_dir, recursive)
    total_files = len(target_files['word']) + len(target_files['excel'])
    
    if total_files == 0:
        logger.warning("没有找到可转换的文件")
        return {'success': 0, 'failed': 0, 'skipped': 0}
    
    # 确定输出目录
    if output_dir:
        output_path = Path(output_dir)
    else:
        output_path = Path(input_dir) / "markdown_output"
    
    # 创建转换器
    try:
        converter = ComprehensiveToMarkdownConverter(
            enable_ocr=enable_ocr,
            enable_table_structure=enable_table_structure
        )
    except Exception as e:
        logger.error(f"创建转换器失败: {str(e)}")
        return {'success': 0, 'failed': 0, 'skipped': 0}
    
    logger.info(f"开始批量转换，输出目录: {output_path}")
    logger.info(f"OCR启用: {enable_ocr}, 表格结构识别启用: {enable_table_structure}")
    
    results = {'success': 0, 'failed': 0, 'skipped': 0}
    current_file = 0
    
    try:
        # 转换Word文件
        for word_file in target_files['word']:
            current_file += 1
            logger.info(f"处理进度: {current_file}/{total_files}")
            
            try:
                # 保持相对路径结构
                relative_path = word_file.relative_to(Path(input_dir))
                output_subdir = output_path / relative_path.parent
                
                # 检查输出文件是否已存在
                markdown_file = output_subdir / f"{word_file.stem}.md"
                if markdown_file.exists() and not overwrite:
                    logger.warning(f"Markdown文件已存在，跳过: {markdown_file}")
                    results['skipped'] += 1
                    continue
                
                # 转换Word文件
                if converter.convert_word_to_markdown(word_file, output_subdir):
                    results['success'] += 1
                else:
                    results['failed'] += 1
                    
            except Exception as e:
                logger.error(f"处理Word文件时出错 {word_file}: {str(e)}")
                results['failed'] += 1
        
        # 转换Excel文件
        for excel_file in target_files['excel']:
            current_file += 1
            logger.info(f"处理进度: {current_file}/{total_files}")
            
            try:
                # 保持相对路径结构
                relative_path = excel_file.relative_to(Path(input_dir))
                output_subdir = output_path / relative_path.parent
                
                # 对于Excel文件，检查是否所有相关的Markdown文件都已存在
                if not overwrite:
                    try:
                        sheet_names = converter._get_excel_sheet_names(excel_file)
                        all_files_exist = True
                        for sheet_name in sheet_names:
                            safe_sheet_name = converter._sanitize_filename(sheet_name)
                            markdown_filename = f"{excel_file.stem}-{safe_sheet_name}.md"
                            markdown_file = output_subdir / markdown_filename
                            if not markdown_file.exists():
                                all_files_exist = False
                                break
                        
                        if all_files_exist:
                            logger.warning(f"Excel文件的所有sheet都已转换，跳过: {excel_file.name}")
                            results['skipped'] += 1
                            continue
                    except Exception as e:
                        logger.warning(f"检查Excel文件sheet时出错，继续转换: {str(e)}")
                
                # 转换Excel文件
                if converter.convert_excel_to_markdown(excel_file, output_subdir):
                    results['success'] += 1
                else:
                    results['failed'] += 1
                    
            except Exception as e:
                logger.error(f"处理Excel文件时出错 {excel_file}: {str(e)}")
                results['failed'] += 1
    
    finally:
        # 清理临时文件
        converter.cleanup()
    
    return results

def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description='综合文档转Markdown转换器 (Word + Excel)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
功能说明:
  - Word文件(.doc/.docx): 先转换为PDF，再转换为Markdown
  - Excel文件(.xlsx/.xls): 每个sheet转换为单独的Markdown文件，文件名格式为"Excel文件名-sheet名.md"

使用示例:
  python comprehensive_to_markdown_converter.py "/Users/oasis/Documents/work/运营知识库"
  python comprehensive_to_markdown_converter.py "/path/to/files" -o "/path/to/output"
  python comprehensive_to_markdown_converter.py "/path/to/files" --no-recursive --overwrite

Excel处理示例:
  输入文件: data.xlsx (包含3个sheet: Sheet1, Sheet2, 用户信息)
  输出文件: 
    - data-Sheet1.md
    - data-Sheet2.md  
    - data-用户信息.md
        """
    )
    
    parser.add_argument(
        'input_dir',
        help='包含Word和Excel文档的输入目录'
    )
    
    parser.add_argument(
        '-o', '--output-dir',
        help='Markdown输出目录（默认为输入目录下的markdown_output文件夹）',
        default=None
    )
    
    parser.add_argument(
        '--no-recursive',
        action='store_true',
        help='不递归搜索子目录中的文件'
    )
    
    parser.add_argument(
        '--overwrite',
        action='store_true',
        help='覆盖已存在的Markdown文件'
    )
    
    parser.add_argument(
        '--no-ocr',
        action='store_true',
        help='禁用OCR功能'
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
    logger.info("综合文档转Markdown转换器 (Word + Excel)")
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
    results = batch_convert_all(
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