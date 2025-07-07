# docs_parse

综合文档转Markdown转换器，支持Word和Excel文件的批量转换。

## 功能特性

### Word文件处理
- 支持 `.doc` 和 `.docx` 格式
- 使用 LibreOffice 将 Word 文件转换为 PDF
- 使用 Docling 库将 PDF 转换为 Markdown
- 保持原始文件夹结构

### Excel文件处理 ⭐ 新功能
- 支持 `.xls` 和 `.xlsx` 格式
- **每个sheet转换为单独的Markdown文件**
- 文件名格式：`Excel文件名-sheet名.md`
- 自动处理包含特殊字符的sheet名称
- 转换为标准的Markdown表格格式

## 安装依赖

```bash
# 创建虚拟环境
python3 -m venv venv_docling
source venv_docling/bin/activate  # Linux/macOS
# 或 Windows: venv_docling\Scripts\activate

# 安装依赖
pip install docling pandas openpyxl

# 安装LibreOffice（用于Word文件转换）
# macOS: brew install --cask libreoffice
# Ubuntu: sudo apt-get install libreoffice
```

## 使用方法

### 基本用法

```bash
# 转换指定目录下的所有Word和Excel文件
python comprehensive_to_markdown_converter.py "/path/to/documents"

# 指定输出目录
python comprehensive_to_markdown_converter.py "/path/to/documents" -o "/path/to/output"

# 不递归搜索子目录
python comprehensive_to_markdown_converter.py "/path/to/documents" --no-recursive

# 覆盖已存在的文件
python comprehensive_to_markdown_converter.py "/path/to/documents" --overwrite
```

### Excel处理示例

假设有一个Excel文件 `data.xlsx`，包含3个sheet：
- Sheet1
- Sheet2  
- 用户信息

转换后会生成以下Markdown文件：
- `data-Sheet1.md`
- `data-Sheet2.md`
- `data-用户信息.md`

每个文件包含对应sheet的表格数据，格式为标准Markdown表格。

## 高级选项

```bash
# 禁用OCR功能（提高处理速度）
python comprehensive_to_markdown_converter.py "/path/to/documents" --no-ocr

# 禁用表格结构识别
python comprehensive_to_markdown_converter.py "/path/to/documents" --no-table-structure

# 组合使用多个选项
python comprehensive_to_markdown_converter.py "/path/to/documents" \
    -o "/output/directory" \
    --overwrite \
    --no-recursive \
    --no-ocr
```

## 文件结构

```
docs_parse/
├── comprehensive_to_markdown_converter.py  # 主转换器脚本
├── pdf_to_markdown_converter.py           # PDF专用转换器
├── word_to_pdf_converter.py               # Word转PDF转换器
├── test_docling.py                        # 测试脚本
├── requirements_docling.txt               # 依赖列表
├── venv_docling/                          # 虚拟环境
└── docs/                                  # 文档目录
```

## 日志文件

转换过程中会生成详细的日志文件：
- `comprehensive_conversion.log` - 综合转换日志
- `pdf_to_markdown_conversion.log` - PDF转换日志
- `word_to_pdf_conversion.log` - Word转换日志

## 注意事项

1. **首次运行**：Docling库首次运行时会下载AI模型文件，可能需要几分钟时间
2. **LibreOffice**：Word文件转换需要安装LibreOffice
3. **内存使用**：处理大型Excel文件时可能需要较多内存
4. **文件名安全**：Excel sheet名称中的特殊字符会被替换为下划线

## 技术栈

- **Docling**: IBM开源的文档处理库，支持高质量的PDF解析
- **Pandas**: 用于Excel文件读取和数据处理
- **OpenPyXL**: 用于获取Excel sheet信息
- **LibreOffice**: 用于Word文件转PDF