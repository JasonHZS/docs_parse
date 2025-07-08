#!/usr/bin/env python3
"""
测试Excel转换功能
"""

import pandas as pd
from pathlib import Path
import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from comprehensive_to_markdown_converter import ComprehensiveToMarkdownConverter

def create_test_excel():
    """创建测试Excel文件"""
    # 创建第一个sheet的数据
    df1 = pd.DataFrame({
        'A': [1, 2, 3, 4],
        'B': ['a', 'b', 'c', 'd'],
        'C': [10.5, 20.5, 30.5, 40.5]
    })

    # 创建第二个sheet的数据
    df2 = pd.DataFrame({
        'X': [10, 20, 30],
        'Y': ['x', 'y', 'z'],
        'Z': [100, 200, 300]
    })

    # 创建第三个sheet的数据（包含特殊字符的sheet名）
    df3 = pd.DataFrame({
        'Name': ['张三', '李四', '王五'],
        'Age': [25, 30, 35],
        'City': ['北京', '上海', '广州']
    })

    # 保存到Excel文件
    with pd.ExcelWriter('test_excel.xlsx') as writer:
        df1.to_excel(writer, sheet_name='Sheet1', index=False)
        df2.to_excel(writer, sheet_name='Sheet2', index=False)
        df3.to_excel(writer, sheet_name='用户信息', index=False)

    print('测试Excel文件创建成功: test_excel.xlsx')
    print('包含3个sheet: Sheet1, Sheet2, 用户信息')

def test_excel_conversion():
    """测试Excel转换功能"""
    # 创建测试文件
    create_test_excel()
    
    # 创建转换器
    converter = ComprehensiveToMarkdownConverter()
    
    # 转换Excel文件
    excel_file = Path('test_excel.xlsx')
    output_dir = Path('test_output')
    
    print(f"\n开始转换Excel文件: {excel_file}")
    success = converter.convert_excel_to_markdown(excel_file, output_dir)
    
    if success:
        print("✅ Excel转换成功!")
        
        # 列出生成的文件
        print("\n生成的文件:")
        for md_file in output_dir.glob('*.md'):
            print(f"  - {md_file.name}")
            
        # 显示第一个文件的内容
        first_file = next(output_dir.glob('*.md'))
        print(f"\n{first_file.name} 的内容:")
        with open(first_file, 'r', encoding='utf-8') as f:
            print(f.read())
    else:
        print("❌ Excel转换失败!")
    
    # 清理
    converter.cleanup()

if __name__ == '__main__':
    test_excel_conversion() 