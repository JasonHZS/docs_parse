# 快速入门指南

## 🚀 立即开始使用

### 1. 激活虚拟环境

```bash
cd /Users/oasis/Documents/GitHub\ Project/docs_parse
source venv_docling/bin/activate
```

### 2. 验证安装

```bash
python test_docling.py
```

### 3. 转换您的 PDF 文件

#### 基本用法
```bash
# 转换当前目录下的所有 PDF 文件
python pdf_to·_markdown_converter.py .

# 转换指定文件夹
python pdf_to_markdown_converter.py /path/to/your/pdf/folder
```

#### 常用选项
```bash
# 指定输出目录
python pdf_to_markdown_converter.py /Users/oasis/Documents/work/增长黑客知识库/pdf -o /Users/oasis/Documents/work/增长黑客知识库/markdown

# 不递归搜索子文件夹
python pdf_to_markdown_converter.py ./pdfs --no-recursive

# 覆盖已存在的文件
python pdf_to_markdown_converter.py ./pdfs --overwrite
```

## 📁 文件结构

创建的文件：
- `pdf_to_markdown_converter.py` - 主转换器脚本
- `test_docling.py` - 测试脚本
- `README_pdf_to_markdown.md` - 详细使用说明
- `QUICK_START.md` - 快速入门指南（本文件）

## ⚡ 快速示例

假设您有一个包含 PDF 文件的文件夹：

```
documents/
├── report1.pdf
├── research_paper.**pdf**
└── manual.pdf
```

运行：
```bash
python pdf_to_markdown_converter.py documents
```

结果：
```
documents/
├── report1.pdf
├── research_paper.pdf  
├── manual.pdf
└── markdown_output/
    ├── report1.md
    ├── research_paper.md
    └── manual.md
```

## 🔧 常见问题

**Q: 首次运行很慢？**
A: 这是正常的，docling 需要下载 AI 模型文件，只需要下载一次。

**Q: 扫描版 PDF 效果不好？**
A: 确保启用了 OCR 功能（默认启用）。

**Q: 表格格式不正确？**
A: 确保启用了 table structure 功能（默认启用）。

## 📖 更多信息

查看详细文档：`README_pdf_to_markdown.md` 