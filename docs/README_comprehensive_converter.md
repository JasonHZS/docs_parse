# 综合文档转Markdown转换器使用指南

## 概述

这个工具可以将指定目录下的Word文档和Excel文件批量转换为Markdown格式：

- **Word文件** (.doc/.docx)：先转换为PDF，再转换为Markdown
- **Excel文件** (.xls/.xlsx)：直接转换为Markdown

## 功能特点

✅ 支持批量转换多种文档格式  
✅ 自动激活docling虚拟环境  
✅ 保持目录结构  
✅ 详细的转换日志  
✅ OCR识别支持（扫描版文档）  
✅ 表格结构识别  
✅ 错误处理和恢复  

## 系统要求

### 必需组件
- Python 3.8+
- docling虚拟环境 (`venv_docling`)
- docling库

### 可选组件（Word转换需要）
- LibreOffice（用于Word转PDF）
- 安装命令：`brew install --cask libreoffice`

## 快速开始

### 方法1：使用便捷脚本（推荐）

```bash
# 运行转换（会自动激活虚拟环境）
./run_comprehensive_converter.sh

# 带自定义选项
./run_comprehensive_converter.sh --overwrite --no-recursive
```

### 方法2：手动运行

```bash
# 激活虚拟环境
source venv_docling/bin/activate

# 运行转换脚本
python comprehensive_to_markdown_converter.py "/Users/oasis/Documents/work/运营知识库"

# 带自定义输出目录
python comprehensive_to_markdown_converter.py "/Users/oasis/Documents/work/运营知识库" -o "/path/to/output"
```

## 命令行选项

| 选项 | 说明 | 默认值 |
|------|------|--------|
| `input_dir` | 输入目录（必需） | - |
| `-o, --output-dir` | 输出目录 | `输入目录/markdown_output` |
| `--overwrite` | 覆盖已存在的文件 | 跳过已存在文件 |
| `--no-recursive` | 不递归搜索子目录 | 递归搜索 |
| `--no-ocr` | 禁用OCR功能 | 启用OCR |
| `--no-table-structure` | 禁用表格结构识别 | 启用表格识别 |

## 使用示例

### 基本转换
```bash
python comprehensive_to_markdown_converter.py "/Users/oasis/Documents/work/运营知识库"
```

### 指定输出目录
```bash
python comprehensive_to_markdown_converter.py "/Users/oasis/Documents/work/运营知识库" -o "/Users/oasis/Desktop/markdown_files"
```

### 覆盖已存在文件
```bash
python comprehensive_to_markdown_converter.py "/Users/oasis/Documents/work/运营知识库" --overwrite
```

### 仅处理当前目录（不递归）
```bash
python comprehensive_to_markdown_converter.py "/Users/oasis/Documents/work/运营知识库" --no-recursive
```

### 禁用OCR（加快处理速度）
```bash
python comprehensive_to_markdown_converter.py "/Users/oasis/Documents/work/运营知识库" --no-ocr
```

## 输出结构

转换后的文件将保存在以下结构中：

```
输入目录/
└── markdown_output/
    ├── 子目录1/
    │   ├── 文档1.md
    │   └── 文档2.md
    └── 子目录2/
        ├── 表格1.md
        └── 表格2.md
```

## 转换流程

### Word文件转换流程
1. **扫描文件** → 识别.doc/.docx文件
2. **Word转PDF** → 使用LibreOffice转换
3. **PDF转Markdown** → 使用docling转换
4. **保存结果** → 清理临时文件

### Excel文件转换流程
1. **扫描文件** → 识别.xls/.xlsx文件
2. **Excel转Markdown** → 使用docling直接转换
3. **保存结果** → 保持表格结构

## 日志文件

转换过程会生成详细日志：

- **文件名**：`comprehensive_conversion.log`
- **内容**：包含所有转换步骤、错误信息和统计数据
- **位置**：当前工作目录

## 故障排除

### 常见问题

#### 1. LibreOffice未安装
```
错误: 未检测到LibreOffice，Word文件转换可能失败
解决: brew install --cask libreoffice
```

#### 2. docling库未安装
```
错误: 无法导入docling库
解决: source venv_docling/bin/activate && pip install docling
```

#### 3. 虚拟环境不存在
```
错误: 找不到docling虚拟环境目录
解决: python -m venv venv_docling
```

#### 4. 转换失败
- 检查文件权限
- 确认文件未损坏
- 查看日志文件了解详细错误

### 性能优化

#### 加快转换速度
- 使用`--no-ocr`禁用OCR（如果不需要处理扫描版文档）
- 使用`--no-table-structure`禁用表格识别（如果不需要表格结构）
- 确保有足够的磁盘空间用于临时文件

#### 处理大量文件
- 分批处理大目录
- 使用`--no-recursive`仅处理当前目录
- 定期清理日志文件

## 技术细节

### 支持的文件格式

#### Word文档
- `.doc` - Microsoft Word 97-2003
- `.docx` - Microsoft Word 2007+

#### Excel表格
- `.xls` - Microsoft Excel 97-2003
- `.xlsx` - Microsoft Excel 2007+

### 转换质量

#### Word文档
- ✅ 文本内容
- ✅ 标题层级
- ✅ 列表和编号
- ✅ 表格结构
- ✅ 图片（转换为引用）
- ⚠️ 复杂格式可能简化

#### Excel表格
- ✅ 表格数据
- ✅ 单元格内容
- ✅ 基本格式
- ⚠️ 公式结果（不保留公式）
- ⚠️ 图表（可能丢失）

## 更新日志

### v1.0.0
- 初始版本
- 支持Word和Excel转换
- 集成docling转换引擎
- 自动虚拟环境管理

## 许可证

本项目遵循MIT许可证。

## 技术支持

如遇问题，请查看：
1. 日志文件 `comprehensive_conversion.log`
2. 确认所有依赖已正确安装
3. 检查文件权限和磁盘空间 