#!/usr/bin/env python3
"""
Word文档批量转PDF转换器 (增强版)
支持.doc和.docx格式的Word文档批量转换为PDF
支持多种转换方法：LibreOffice、pandoc、docx2pdf
"""

import os
import sys
import argparse
import subprocess
import shutil
from pathlib import Path
from typing import List, Optional
import logging

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('word_to_pdf_conversion.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class WordToPdfConverter:
    """Word转PDF转换器类"""
    
    def __init__(self):
        self.available_methods = self._check_available_methods()
        self.preferred_method = self._get_preferred_method()
        
    def _check_available_methods(self) -> dict:
        """检查可用的转换方法"""
        methods = {
            'libreoffice': False,
            'pandoc': False,
            'docx2pdf': False
        }
        
        # 检查LibreOffice
        if shutil.which('libreoffice') or shutil.which('soffice'):
            methods['libreoffice'] = True
            logger.info("检测到LibreOffice")
        
        # 检查pandoc
        if shutil.which('pandoc'):
            methods['pandoc'] = True
            logger.info("检测到pandoc")
        
        # 检查docx2pdf
        try:
            import docx2pdf
            methods['docx2pdf'] = True
            logger.info("检测到docx2pdf")
        except ImportError:
            pass
        
        return methods
    
    def _get_preferred_method(self) -> Optional[str]:
        """获取首选的转换方法"""
        # 优先级：LibreOffice > pandoc > docx2pdf
        if self.available_methods['libreoffice']:
            return 'libreoffice'
        elif self.available_methods['pandoc']:
            return 'pandoc'
        elif self.available_methods['docx2pdf']:
            return 'docx2pdf'
        else:
            return None
    
    def _convert_with_libreoffice(self, word_file: Path, output_dir: Path) -> bool:
        """使用LibreOffice进行转换"""
        try:
            # 使用LibreOffice的无头模式转换
            # 优先使用soffice命令
            if shutil.which('soffice'):
                cmd_name = 'soffice'
            else:
                cmd_name = 'libreoffice'
            
            cmd = [
                cmd_name,
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', str(output_dir),
                str(word_file)
            ]
            
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=60  # 60秒超时
            )
            
            if result.returncode == 0:
                pdf_file = output_dir / f"{word_file.stem}.pdf"
                if pdf_file.exists() and pdf_file.stat().st_size > 0:
                    return True
            
            logger.error(f"LibreOffice转换失败: {result.stderr}")
            return False
            
        except subprocess.TimeoutExpired:
            logger.error(f"LibreOffice转换超时: {word_file.name}")
            return False
        except Exception as e:
            logger.error(f"LibreOffice转换出错: {str(e)}")
            return False
    
    def convert_word_to_pdf(self, word_file: Path, output_dir: Path, method: Optional[str] = None) -> bool:
        """
        将Word文档转换为PDF
        
        Args:
            word_file: Word文档路径
            output_dir: 输出目录
            method: 指定转换方法
            
        Returns:
            转换是否成功
        """
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # 检查PDF文件是否已存在
        pdf_file = output_dir / f"{word_file.stem}.pdf"
        if pdf_file.exists():
            logger.warning(f"PDF文件已存在，跳过: {pdf_file}")
            return True
        
        # 选择转换方法
        if method and method in self.available_methods and self.available_methods[method]:
            conversion_method = method
        else:
            conversion_method = self.preferred_method
        
        if not conversion_method:
            logger.error("没有可用的转换方法！")
            return False
        
        logger.info(f"使用 {conversion_method} 转换: {word_file.name} -> {pdf_file.name}")
        
        # 执行转换
        if conversion_method == 'libreoffice':
            return self._convert_with_libreoffice(word_file, output_dir)
        
        return False

def find_word_files(directory: str) -> List[Path]:
    """查找指定目录下的所有Word文档"""
    directory_path = Path(directory)
    if not directory_path.exists():
        logger.error(f"目录不存在: {directory}")
        return []
    
    if not directory_path.is_dir():
        logger.error(f"路径不是目录: {directory}")
        return []
    
    word_extensions = ['*.doc', '*.docx']
    word_files = []
    
    for extension in word_extensions:
        files = list(directory_path.rglob(extension))
        word_files.extend(files)
    
    word_files = sorted(list(set(word_files)))
    logger.info(f"在目录 {directory} 中找到 {len(word_files)} 个Word文档")
    
    return word_files

def batch_convert(input_dir: str, output_dir: Optional[str] = None, 
                 skip_existing: bool = True, method: Optional[str] = None) -> dict:
    """批量转换Word文档为PDF"""
    
    converter = WordToPdfConverter()
    
    # 检查是否有可用的转换方法
    if not converter.preferred_method:
        logger.error("没有找到可用的转换工具！")
        logger.error("请安装LibreOffice: brew install --cask libreoffice")
        return {'success': 0, 'failed': 0, 'skipped': 0}
    
    logger.info(f"使用转换方法: {converter.preferred_method}")
    logger.info(f"可用方法: {[k for k, v in converter.available_methods.items() if v]}")
    
    word_files = find_word_files(input_dir)
    if not word_files:
        logger.warning("没有找到Word文档")
        return {'success': 0, 'failed': 0, 'skipped': 0}
    
    # 确定输出目录
    if output_dir:
        output_path = Path(output_dir)
    else:
        output_path = Path(input_dir)
    
    results = {'success': 0, 'failed': 0, 'skipped': 0}
    
    for i, word_file in enumerate(word_files, 1):
        logger.info(f"处理进度: {i}/{len(word_files)}")
        
        try:
            # 检查是否跳过已存在的PDF
            pdf_file = output_path / f"{word_file.stem}.pdf"
            if skip_existing and pdf_file.exists():
                logger.info(f"PDF已存在，跳过: {pdf_file}")
                results['skipped'] += 1
                continue
            
            # 转换文档
            if converter.convert_word_to_pdf(word_file, output_path, method):
                results['success'] += 1
                logger.info(f"转换成功: {word_file.name}")
            else:
                results['failed'] += 1
                logger.error(f"转换失败: {word_file.name}")
                
        except Exception as e:
            logger.error(f"处理文件时出错 {word_file}: {str(e)}")
            results['failed'] += 1
    
    return results

def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description='批量将Word文档转换为PDF (增强版)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  python word_to_pdf_converter_enhanced.py /path/to/word/files
  python word_to_pdf_converter_enhanced.py /path/to/word/files -o /path/to/output
  python word_to_pdf_converter_enhanced.py --check  # 检查可用方法
        """
    )
    
    parser.add_argument(
        'input_dir',
        nargs='?',  # 使input_dir成为可选参数
        help='包含Word文档的输入目录'
    )
    
    parser.add_argument(
        '-o', '--output-dir',
        help='PDF输出目录（默认为原文件所在目录）',
        default=None
    )
    
    parser.add_argument(
        '--force',
        action='store_true',
        help='强制覆盖已存在的PDF文件'
    )
    
    parser.add_argument(
        '--check',
        action='store_true',
        help='仅检查可用的转换方法'
    )
    
    args = parser.parse_args()
    
    # 如果只是检查可用方法
    if args.check:
        converter = WordToPdfConverter()
        print("可用的转换方法:")
        for method, available in converter.available_methods.items():
            status = "✓" if available else "✗"
            print(f"  {status} {method}")
        
        if converter.preferred_method:
            print(f"\n首选方法: {converter.preferred_method}")
        else:
            print("\n没有可用的转换方法！")
            print("请安装LibreOffice: brew install --cask libreoffice")
        
        return
    
    # 如果不是检查模式，则需要input_dir参数
    if not args.input_dir:
        parser.error("需要提供输入目录参数，或使用 --check 检查可用方法")
    
    # 验证输入目录
    if not os.path.exists(args.input_dir):
        logger.error(f"输入目录不存在: {args.input_dir}")
        sys.exit(1)
    
    if not os.path.isdir(args.input_dir):
        logger.error(f"输入路径不是目录: {args.input_dir}")
        sys.exit(1)
    
    logger.info("开始批量转换Word文档为PDF")
    logger.info(f"输入目录: {args.input_dir}")
    
    if args.output_dir:
        logger.info(f"输出目录: {args.output_dir}")
    else:
        logger.info("输出目录: 原文件所在目录")
    
    # 执行批量转换
    results = batch_convert(
        input_dir=args.input_dir,
        output_dir=args.output_dir,
        skip_existing=not args.force
    )
    
    # 输出结果统计
    logger.info("转换完成!")
    logger.info(f"成功: {results['success']} 个文件")
    logger.info(f"失败: {results['failed']} 个文件")
    logger.info(f"跳过: {results['skipped']} 个文件")
    
    if results['failed'] > 0:
        logger.warning("部分文件转换失败，请查看日志了解详情")
        sys.exit(1)

if __name__ == '__main__':
    main() 