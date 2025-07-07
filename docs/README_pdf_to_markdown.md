# PDF 转 Markdown 转换器

这是一个基于 [Docling](https://github.com/docling-project/docling) 库的 PDF 批量转 Markdown 工具。Docling 是 IBM 开源的文档处理库，支持高质量的 PDF 解析，包括表格结构识别、公式提取、OCR 等功能。

## 特性

- 🗂️ 批量转换文件夹中的所有 PDF 文件
- 📑 先进的 PDF 理解能力，包括页面布局、阅读顺序、表格结构等
- 🔍 OCR 支持，可处理扫描版 PDF
- 📊 表格结构识别和保持
- 🧮 公式和代码块识别
- 📁 保持原始文件夹结构
- 📝 详细的转换日志
- ⚙️ 灵活的配置选项

## 安装

### 1. 创建虚拟环境

```bash
python3 -m venv venv_docling
source venv_docling/bin/activate  # Linux/macOS
# 或者 Windows:
# venv_docling\Scripts\activate
```

### 2. 安装 docling

使用 uv（推荐）：
```bash
uv pip install docling
```

或使用 pip：
```bash
pip install docling
```

## 使用方法

### 基本用法

转换指定文件夹中的所有 PDF 文件：

```bash
python pdf_to_markdown_converter.py /path/to/pdf/folder
```

### 高级用法

```bash
# 指定输出目录
python pdf_to_markdown_converter.py /path/to/pdf/folder -o /path/to/output

# 不递归搜索子文件夹
python pdf_to_markdown_converter.py /path/to/pdf/folder --no-recursive

# 覆盖已存在的 Markdown 文件
python pdf_to_markdown_converter.py /path/to/pdf/folder --overwrite

# 禁用 OCR（提高处理速度，但扫描版 PDF 效果较差）
python pdf_to_markdown_converter.py /path/to/pdf/folder --no-ocr

# 禁用表格结构识别
python pdf_to_markdown_converter.py /path/to/pdf/folder --no-table-structure

# 组合使用多个选项
python pdf_to_markdown_converter.py /path/to/pdf/folder \
    -o /output/directory \
    --overwrite \
    --no-recursive \
    --no-ocr
```

### 参数说明

- `input_dir`: 包含 PDF 文件的输入目录（必需）
- `-o, --output-dir`: Markdown 输出目录（可选，默认为输入目录下的 `markdown_output` 文件夹）
- `--no-recursive`: 不递归搜索子目录中的 PDF 文件
- `--overwrite`: 覆盖已存在的 Markdown 文件
- `--no-ocr`: 禁用 OCR 功能
- `--no-table-structure`: 禁用表格结构识别

## 输出说明

### 文件结构

转换后会保持原始的文件夹结构：

```
输入目录/
├── document1.pdf
├── subfolder/
│   └── document2.pdf
└── document3.pdf

输出目录/
├── document1.md
├── subfolder/
│   └── document2.md
└── document3.md
```

### 转换日志

程序会生成详细的转换日志：
- 控制台输出：实时显示转换进度
- 日志文件：`pdf_to_markdown_conversion.log`

## 性能和质量

### Docling 的优势

1. **高质量转换**：相比其他 PDF 解析工具，Docling 提供更准确的文档结构识别
2. **表格处理**：能够准确识别和保持复杂表格结构
3. **OCR 能力**：内置 OCR 引擎，支持扫描版 PDF
4. **多模态理解**：支持图片、公式、代码块等内容的识别

### 性能建议

- **启用 OCR**：对于扫描版 PDF 必需，但会增加处理时间
- **表格识别**：对于包含大量表格的文档建议保持启用
- **批量处理**：大量文件时建议分批处理，避免内存占用过高

## 故障排除

### 常见问题

1. **导入错误**
   ```
   错误：无法导入docling库
   ```
   解决：确保已正确安装 docling：`uv pip install docling`

2. **内存不足**
   - 减少并发处理的文件数量
   - 禁用不必要的功能（如 OCR）
   - 增加系统内存

3. **转换失败**
   - 检查 PDF 文件是否损坏
   - 查看日志文件了解具体错误信息
   - 对于复杂 PDF，可尝试禁用某些功能

### 支持的 PDF 类型

- ✅ 文本型 PDF
- ✅ 扫描版 PDF（需启用 OCR）
- ✅ 混合型 PDF（文本+图像）
- ✅ 包含表格的 PDF
- ✅ 包含公式的 PDF
- ⚠️ 加密 PDF（需先解密）

## 示例

### 简单示例

```bash
# 转换单个文件夹中的 PDF
python pdf_to_markdown_converter.py ./documents

# 转换结果会保存在 ./documents/markdown_output/ 中
```

### 复杂示例

```bash
# 转换整个文档库，保持文件夹结构
python pdf_to_markdown_converter.py /Users/documents/research_papers \
    -o /Users/documents/markdown_papers \
    --overwrite
```

## 注意事项

1. **第一次使用**：Docling 会下载必需的模型文件，可能需要几分钟时间
2. **网络连接**：首次运行需要网络连接来下载模型
3. **磁盘空间**：确保有足够的磁盘空间存储输出文件
4. **文件编码**：输出的 Markdown 文件使用 UTF-8 编码

## 技术细节

### Docling 配置

脚本使用以下 Docling 配置：
- OCR：EasyOCR（默认）
- 表格识别：启用单元格匹配
- 公式识别：自动检测
- 图像分类：自动识别

### 性能调优

可以通过环境变量调整性能：
```bash
# 限制 CPU 线程数
export OMP_NUM_THREADS=4

# 运行转换
python pdf_to_markdown_converter.py /path/to/pdfs
```

## 更多信息

- [Docling 官方文档](https://docling-project.github.io/docling/)
- [Docling GitHub 仓库](https://github.com/docling-project/docling)
- [Docling 技术报告](https://arxiv.org/abs/2408.09869) 