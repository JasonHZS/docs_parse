#!/usr/bin/env python3
"""
Enhanced Excel to Markdown Converter
针对RAG检索优化的Excel转换器，提升可读性和语义结构

新增功能：键值对字符串格式
- 将表格数据转换为 "列名：值；列名：值" 的格式
- 每行数据组合成一个完整的字符串，便于RAG检索

使用示例：
    # 1. 使用便捷函数（推荐）- 使用sheet名称作为后缀
    from improved_excel_to_markdown import convert_excel_to_key_value_strings
    
    success = convert_excel_to_key_value_strings(
        excel_path="/path/to/your/file.xlsx",
        output_path="/path/to/output/directory"
    )
    
    # 2. 使用便捷函数 - 使用自定义后缀
    from improved_excel_to_markdown import convert_excel_to_key_value_strings_with_custom_suffix
    
    success = convert_excel_to_key_value_strings_with_custom_suffix(
        excel_path="/path/to/your/file.xlsx",
        output_path="/path/to/output/directory",
        custom_suffix="—数据"
    )
    
    # 3. 使用类实例
    converter = ImprovedExcelToMarkdownConverter(
        output_format="key_value_strings"
    )
    converter.convert_excel_file(excel_file, output_dir)

输出格式示例：
    原始Excel表格（假设sheet名称为"员工信息"）：
    | 姓名 | 年龄 | 性别 |
    | 张三 | 18   | 男   |
    | 李四 | 20   | 女   |
    
    转换后的Markdown：
    姓名：张三；年龄：18；性别：男 —员工信息
    姓名：李四；年龄：20；性别：女 —员工信息

支持的输出格式：
- "key_value_strings": 键值对字符串格式 (新增)
- "table": 标准表格格式
- "form": 表单格式
- "sections": 分节格式  
- "auto": 自动选择格式 (默认)
"""

import os
import re
import pandas as pd
import openpyxl
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple
import logging
import argparse

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class ImprovedExcelToMarkdownConverter:
    """优化的Excel到Markdown转换器"""
    
    def __init__(self, output_format: str = "auto", suffix: str = "—信息"):
        """
        初始化转换器
        
        Args:
            output_format: 输出格式 ('auto', 'table', 'form', 'sections', 'key_value_strings')
            suffix: 键值对字符串格式的后缀
        """
        self.output_format = output_format
        self.suffix = suffix
        
    def detect_table_structure(self, df: pd.DataFrame) -> Dict[str, Any]:
        """智能检测表格结构和类型"""
        structure_info = {
            'has_headers': False,
            'header_row': 0,
            'data_type': 'unknown',
            'key_columns': [],
            'summary_columns': [],
            'is_list_format': False,
            'is_form_format': False
        }
        
        if df.empty:
            return structure_info
            
        # 检测是否有表头
        first_row = df.iloc[0] if len(df) > 0 else pd.Series()
        second_row = df.iloc[1] if len(df) > 1 else pd.Series()
        
        # 如果第一行与后续行数据类型差异较大，可能是表头
        if len(df) > 1:
            first_row_types = [type(x).__name__ for x in first_row.fillna('')]
            second_row_types = [type(x).__name__ for x in second_row.fillna('')]
            
            type_diff_ratio = sum(1 for a, b in zip(first_row_types, second_row_types) if a != b) / len(first_row_types)
            if type_diff_ratio > 0.5 or all(isinstance(x, str) for x in first_row.fillna('')):
                structure_info['has_headers'] = True
                
        # 检测表格类型
        num_cols = len(df.columns)
        num_rows = len(df)
        
        # 表单格式检测（键值对形式）
        if num_cols == 2 and num_rows > 3:
            # 检查是否是键值对格式
            first_col_str_ratio = sum(1 for x in df.iloc[:, 0] if isinstance(x, str)) / num_rows
            if first_col_str_ratio > 0.8:
                structure_info['is_form_format'] = True
                structure_info['data_type'] = 'form'
                
        # 列表格式检测
        elif num_cols >= 3 and num_rows >= 2:
            structure_info['is_list_format'] = True
            structure_info['data_type'] = 'table'
            
        # 检测关键列（包含重要信息的列）
        for i, col in enumerate(df.columns):
            col_data = df.iloc[:, i].fillna('')
            unique_ratio = len(set(col_data)) / len(col_data) if len(col_data) > 0 else 0
            
            # 如果列的唯一值比例高，可能是关键列
            if unique_ratio > 0.7:
                structure_info['key_columns'].append(i)
                
        return structure_info
    
    def clean_cell_content(self, content: Any) -> str:
        """清理单元格内容"""
        if pd.isna(content) or content == '':
            return ''
            
        # 转换为字符串并清理
        text = str(content).strip()
        
        # 移除多余的空白字符
        text = re.sub(r'\s+', ' ', text)
        
        # 保留完整内容，不进行长度限制
            
        # 转义Markdown特殊字符
        text = text.replace('|', '\\|').replace('\n', ' ').replace('\r', ' ')
        
        return text
    
    def format_as_form(self, df: pd.DataFrame, sheet_name: str) -> str:
        """将表格格式化为表单样式（适合键值对数据）"""
        markdown_content = f"# {sheet_name}\n\n"
        
        if len(df.columns) >= 2:
            for idx, row in df.iterrows():
                key = self.clean_cell_content(row.iloc[0])
                value = self.clean_cell_content(row.iloc[1])
                
                if key and value:
                    markdown_content += f"**{key}**: {value}\n\n"
                elif key:
                    markdown_content += f"## {key}\n\n"
                    
        return markdown_content
    
    def format_as_key_value_strings(self, df: pd.DataFrame, sheet_name: str, structure_info: Dict, use_sheet_name_as_suffix: bool = True) -> str:
        """将表格格式化为键值对字符串（列名：值；列名：值的格式）"""
        markdown_content = f"# {sheet_name}\n\n"
        
        # 处理表头
        if structure_info['has_headers']:
            headers = [self.clean_cell_content(df.iloc[0, i]) for i in range(len(df.columns))]
            data_start_row = 1
        else:
            headers = [f"列{i+1}" for i in range(len(df.columns))]
            data_start_row = 0
            
        # 过滤空的表头
        non_empty_cols = [i for i, header in enumerate(headers) if header.strip()]
        
        if not non_empty_cols:
            return f"# {sheet_name}\n\n*此表格无有效数据*\n\n"
            
        filtered_headers = [headers[i] for i in non_empty_cols]
        
        # 确定使用的后缀
        if use_sheet_name_as_suffix:
            suffix = f"—{sheet_name}"
        else:
            suffix = self.suffix
        
        # 生成键值对字符串
        valid_rows = 0
        for idx in range(data_start_row, len(df)):
            row_data = [self.clean_cell_content(df.iloc[idx, i]) for i in non_empty_cols]
            
            # 跳过完全空的行
            if not any(cell.strip() for cell in row_data):
                continue
                
            # 组合键值对
            key_value_pairs = []
            for header, value in zip(filtered_headers, row_data):
                if value.strip():  # 只包含有值的字段
                    key_value_pairs.append(f"{header}：{value}")
                elif header.strip():  # 如果没有值但有表头，显示空值
                    key_value_pairs.append(f"{header}：")
                    
            if key_value_pairs:
                # 用分号连接各个键值对，末尾添加后缀
                key_value_string = "；".join(key_value_pairs)
                if suffix:
                    key_value_string += f" {suffix}"
                markdown_content += f"{key_value_string}\n\n"
                valid_rows += 1
                
        if valid_rows == 0:
            markdown_content += "\n*此表格无有效数据行*\n"
            
        return markdown_content
    
    def format_as_structured_table(self, df: pd.DataFrame, sheet_name: str, structure_info: Dict) -> str:
        """格式化为结构化表格"""
        markdown_content = f"# {sheet_name}\n\n"
        
        # 处理表头
        if structure_info['has_headers']:
            headers = [self.clean_cell_content(df.iloc[0, i]) for i in range(len(df.columns))]
            data_start_row = 1
        else:
            headers = [f"列{i+1}" for i in range(len(df.columns))]
            data_start_row = 0
            
        # 过滤空的表头
        non_empty_cols = [i for i, header in enumerate(headers) if header.strip()]
        
        if not non_empty_cols:
            return f"# {sheet_name}\n\n*此表格无有效数据*\n\n"
            
        # 构建表格
        filtered_headers = [headers[i] for i in non_empty_cols]
        
        # 表格头部
        markdown_content += "| " + " | ".join(filtered_headers) + " |\n"
        markdown_content += "| " + " | ".join(["---"] * len(filtered_headers)) + " |\n"
        
        # 数据行
        valid_rows = 0
        for idx in range(data_start_row, len(df)):
            row_data = [self.clean_cell_content(df.iloc[idx, i]) for i in non_empty_cols]
            
            # 跳过完全空的行
            if not any(cell.strip() for cell in row_data):
                continue
                
            markdown_content += "| " + " | ".join(row_data) + " |\n"
            valid_rows += 1
            
            # 保留所有数据行，不进行数量限制
                
        if valid_rows == 0:
            markdown_content += "\n*此表格无有效数据行*\n"
            
        return markdown_content + "\n"
    
    def format_as_sections(self, df: pd.DataFrame, sheet_name: str) -> str:
        """将复杂表格格式化为分节内容（提升RAG检索效果）"""
        markdown_content = f"# {sheet_name}\n\n"
        
        # 尝试按行分组处理
        current_section = None
        section_content = []
        
        for idx, row in df.iterrows():
            row_data = [self.clean_cell_content(cell) for cell in row]
            
            # 检测是否为新节标题（第一列有内容，其他列大部分为空）
            first_col = row_data[0] if row_data else ''
            other_cols = row_data[1:] if len(row_data) > 1 else []
            
            is_section_header = (
                first_col.strip() and 
                sum(1 for cell in other_cols if cell.strip()) <= 1
            )
            
            if is_section_header and len(first_col) < 50:
                # 保存之前的节
                if current_section and section_content:
                    markdown_content += f"## {current_section}\n\n"
                    for content in section_content:
                        markdown_content += f"- {content}\n"
                    markdown_content += "\n"
                    
                # 开始新节
                current_section = first_col
                section_content = []
                
                # 如果有其他列的内容，也加入
                for cell in other_cols:
                    if cell.strip():
                        section_content.append(cell)
                        
            else:
                # 添加到当前节
                non_empty_cells = [cell for cell in row_data if cell.strip()]
                if non_empty_cells:
                    if len(non_empty_cells) == 1:
                        section_content.append(non_empty_cells[0])
                    else:
                        # 多列内容用连字符连接
                        section_content.append(" - ".join(non_empty_cells))
                        
        # 保存最后一节
        if current_section and section_content:
            markdown_content += f"## {current_section}\n\n"
            for content in section_content:
                markdown_content += f"- {content}\n"
            markdown_content += "\n"
            
        return markdown_content
    
    def convert_sheet_to_markdown(self, excel_file: Path, sheet_name: str) -> Optional[str]:
        """将单个Excel工作表转换为优化的Markdown"""
        try:
            # 读取数据
            df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
            
            if df.empty:
                return f"# {sheet_name}\n\n*此工作表为空*\n\n"
                
            # 移除完全空的行和列
            df = df.dropna(how='all').dropna(axis=1, how='all')
            
            if df.empty:
                return f"# {sheet_name}\n\n*此工作表无有效数据*\n\n"
                
            # 分析表格结构
            structure_info = self.detect_table_structure(df)
            
            # 根据输出格式选择合适的格式化方法
            if self.output_format == "key_value_strings":
                return self.format_as_key_value_strings(df, sheet_name, structure_info, use_sheet_name_as_suffix=True)
            elif self.output_format == "form":
                return self.format_as_form(df, sheet_name)
            elif self.output_format == "table":
                return self.format_as_structured_table(df, sheet_name, structure_info)
            elif self.output_format == "sections":
                return self.format_as_sections(df, sheet_name)
            else:
                # 自动格式选择 (原有逻辑)
                if structure_info['is_form_format']:
                    return self.format_as_form(df, sheet_name)
                elif structure_info['is_list_format'] and len(df.columns) <= 6:
                    return self.format_as_structured_table(df, sheet_name, structure_info)
                else:
                    # 复杂表格使用分节格式
                    return self.format_as_sections(df, sheet_name)
                
        except Exception as e:
            logger.error(f"转换工作表 '{sheet_name}' 时出错: {str(e)}")
            return f"# {sheet_name}\n\n*转换此工作表时出错: {str(e)}*\n\n"
    
    def convert_excel_file(self, excel_file: Path, output_dir: Path) -> bool:
        """转换整个Excel文件"""
        try:
            # 获取所有工作表名称
            workbook = openpyxl.load_workbook(excel_file, read_only=True)
            sheet_names = workbook.sheetnames
            workbook.close()
            
            if not sheet_names:
                logger.warning(f"Excel文件无工作表: {excel_file.name}")
                return False
                
            output_dir.mkdir(parents=True, exist_ok=True)
            success_count = 0
            
            for sheet_name in sheet_names:
                try:
                    # 转换工作表
                    markdown_content = self.convert_sheet_to_markdown(excel_file, sheet_name)
                    
                    if markdown_content:
                        # 生成安全的文件名
                        safe_sheet_name = re.sub(r'[<>:"/\\|?*]', '_', sheet_name).strip()
                        if not safe_sheet_name:
                            safe_sheet_name = f"sheet_{sheet_names.index(sheet_name) + 1}"
                            
                        markdown_filename = f"{excel_file.stem}-{safe_sheet_name}.md"
                        markdown_file = output_dir / markdown_filename
                        
                        # 保存文件
                        with open(markdown_file, 'w', encoding='utf-8') as f:
                            f.write(markdown_content)
                            
                        logger.info(f"转换完成: {sheet_name} -> {markdown_file.name}")
                        success_count += 1
                        
                except Exception as e:
                    logger.error(f"处理工作表 '{sheet_name}' 时出错: {str(e)}")
                    continue
                    
            return success_count > 0
            
        except Exception as e:
            logger.error(f"转换Excel文件失败 {excel_file.name}: {str(e)}")
            return False

def convert_excel_to_key_value_strings(excel_path: str, output_path: str, use_sheet_name_as_suffix: bool = True) -> bool:
    """
    便捷函数：将Excel文件转换为键值对字符串格式
    
    Args:
        excel_path: Excel文件路径
        output_path: 输出目录路径
        use_sheet_name_as_suffix: 是否使用sheet名称作为后缀（默认为True）
        
    Returns:
        bool: 转换是否成功
        
    Example:
        >>> convert_excel_to_key_value_strings("/path/to/file.xlsx", "/path/to/output")
        # 输出格式：姓名：张三；年龄：18；性别：男 —Sheet1
        # 如果Excel有多个sheet，每个sheet的数据会有对应的sheet名称作为后缀
    """
    try:
        excel_file = Path(excel_path)
        output_dir = Path(output_path)
        
        if not excel_file.exists():
            logger.error(f"Excel文件不存在: {excel_path}")
            return False
            
        converter = ImprovedExcelToMarkdownConverter(output_format="key_value_strings")
        result = converter.convert_excel_file(excel_file, output_dir)
        
        if result:
            logger.info(f"转换成功！输出文件位于: {output_dir}")
        else:
            logger.error("转换失败")
            
        return result
        
    except Exception as e:
        logger.error(f"转换过程中出错: {str(e)}")
        return False

def convert_excel_to_key_value_strings_with_custom_suffix(excel_path: str, output_path: str, custom_suffix: str = "—信息") -> bool:
    """
    便捷函数：将Excel文件转换为键值对字符串格式（使用自定义后缀）
    
    Args:
        excel_path: Excel文件路径
        output_path: 输出目录路径
        custom_suffix: 自定义后缀（默认为"—信息"）
        
    Returns:
        bool: 转换是否成功
        
    Example:
        >>> convert_excel_to_key_value_strings_with_custom_suffix("/path/to/file.xlsx", "/path/to/output", "—数据")
        # 输出格式：姓名：张三；年龄：18；性别：男 —数据
    """
    try:
        excel_file = Path(excel_path)
        output_dir = Path(output_path)
        
        if not excel_file.exists():
            logger.error(f"Excel文件不存在: {excel_path}")
            return False
            
        converter = ImprovedExcelToMarkdownConverter(output_format="key_value_strings", suffix=custom_suffix)
        
        # 临时修改转换器的行为，使用自定义后缀
        original_method = converter.format_as_key_value_strings
        def custom_format_method(df, sheet_name, structure_info, use_sheet_name_as_suffix=True):
            return original_method(df, sheet_name, structure_info, use_sheet_name_as_suffix=False)
        converter.format_as_key_value_strings = custom_format_method
        
        result = converter.convert_excel_file(excel_file, output_dir)
        
        if result:
            logger.info(f"转换成功！输出文件位于: {output_dir}")
        else:
            logger.error("转换失败")
            
        return result
        
    except Exception as e:
        logger.error(f"转换过程中出错: {str(e)}")
        return False

def main():
    """主函数 - 支持命令行参数"""
    parser = argparse.ArgumentParser(
        description="Enhanced Excel to Markdown Converter - 将Excel文件转换为优化的Markdown格式",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  # 转换单个Excel文件
  python improved_excel_to_markdown.py -i /path/to/file.xlsx -o /path/to/output
  
  # 转换目录中的所有Excel文件
  python improved_excel_to_markdown.py -i /path/to/excel/dir -o /path/to/output
  
  # 使用键值对字符串格式
  python improved_excel_to_markdown.py -i /path/to/file.xlsx -o /path/to/output -f key_value_strings
  
  # 使用自定义后缀
  python improved_excel_to_markdown.py -i /path/to/file.xlsx -o /path/to/output -f key_value_strings -s "—数据"
  
  # 限制处理文件数量
  python improved_excel_to_markdown.py -i /path/to/excel/dir -o /path/to/output --limit 5
        """
    )
    
    parser.add_argument(
        '-i', '--input', 
        required=True,
        help='输入Excel文件路径或包含Excel文件的目录路径'
    )
    
    parser.add_argument(
        '-o', '--output',
        required=True,
        help='输出目录路径'
    )
    
    parser.add_argument(
        '-f', '--format',
        choices=['auto', 'key_value_strings', 'table', 'form', 'sections'],
        default='key_value_strings',
        help='输出格式 (默认: key_value_strings)'
    )
    
    parser.add_argument(
        '-s', '--suffix',
        default='—信息',
        help='自定义后缀 (默认: —信息)'
    )
    
    parser.add_argument(
        '--use-sheet-name-as-suffix',
        action='store_true',
        default=True,
        help='使用sheet名称作为后缀 (默认: True)'
    )
    
    parser.add_argument(
        '--use-custom-suffix',
        action='store_true',
        default=False,
        help='使用自定义后缀而不是sheet名称'
    )
    
    parser.add_argument(
        '--limit',
        type=int,
        default=None,
        help='限制处理的Excel文件数量 (默认: 处理所有文件)'
    )
    
    parser.add_argument(
        '--verbose', '-v',
        action='store_true',
        help='显示详细输出'
    )
    
    args = parser.parse_args()
    
    # 设置日志级别
    if args.verbose:
        logging.basicConfig(level=logging.DEBUG)
    else:
        logging.basicConfig(level=logging.INFO)
    
    # 验证输入路径
    input_path = Path(args.input)
    if not input_path.exists():
        logger.error(f"输入路径不存在: {input_path}")
        return
    
    # 创建输出目录
    output_path = Path(args.output)
    output_path.mkdir(parents=True, exist_ok=True)
    
    # 确定后缀使用方式
    use_sheet_name_as_suffix = not args.use_custom_suffix
    
    # 创建转换器实例
    converter = ImprovedExcelToMarkdownConverter(
        output_format=args.format,
        suffix=args.suffix
    )
    
    # 如果需要使用自定义后缀，修改格式化方法
    if args.use_custom_suffix and args.format == "key_value_strings":
        original_method = converter.format_as_key_value_strings
        def custom_format_method(df, sheet_name, structure_info, use_sheet_name_as_suffix=True):
            return original_method(df, sheet_name, structure_info, use_sheet_name_as_suffix=False)
        converter.format_as_key_value_strings = custom_format_method
    
    # 获取Excel文件列表
    excel_files = []
    
    if input_path.is_file():
        # 单个文件
        if input_path.suffix.lower() in ['.xlsx', '.xls']:
            excel_files.append(input_path)
        else:
            logger.error(f"不支持的文件格式: {input_path.suffix}")
            return
    else:
        # 目录
        excel_files = list(input_path.glob("*.xlsx")) + list(input_path.glob("*.xls"))
    
    if not excel_files:
        logger.info("未找到Excel文件")
        return
    
    # 应用文件数量限制
    if args.limit:
        excel_files = excel_files[:args.limit]
    
    logger.info(f"找到 {len(excel_files)} 个Excel文件")
    
    # 处理文件
    success_count = 0
    for excel_file in excel_files:
        logger.info(f"处理文件: {excel_file.name}")
        try:
            result = converter.convert_excel_file(excel_file, output_path)
            if result:
                success_count += 1
        except Exception as e:
            logger.error(f"处理文件 {excel_file.name} 时出错: {str(e)}")
            continue
    
    # 输出结果统计
    logger.info(f"转换完成！成功处理 {success_count}/{len(excel_files)} 个文件")
    logger.info(f"输出目录: {output_path}")

if __name__ == "__main__":
    main()