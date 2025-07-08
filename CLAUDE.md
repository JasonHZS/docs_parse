# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a comprehensive document conversion tool that transforms Word and Excel files into Markdown format. The project focuses on batch conversion with support for multiple document types and formats.

### Core Architecture

- **comprehensive_to_markdown_converter.py**: Main conversion script that handles both Word and Excel files
- **pdf_to_markdown_converter.py**: Specialized PDF to Markdown converter using Docling
- **word_to_pdf_converter.py**: Word to PDF converter using LibreOffice
- **test_docling.py** and **test_excel_conversion.py**: Test scripts for validation

### Key Technologies

- **Docling**: IBM's document processing library for high-quality PDF parsing and conversion
- **Pandas + OpenPyXL**: Excel file processing (each sheet becomes a separate Markdown file)
- **LibreOffice**: Word to PDF conversion (headless mode)
- **Python 3.9+**: Core runtime in virtual environment

## Common Commands

### Environment Setup
```bash
# Create and activate virtual environment
python3 -m venv venv_docling
source venv_docling/bin/activate

# Install dependencies
pip install -r requirements_docling.txt

# Install LibreOffice (required for Word conversion)
# macOS: brew install --cask libreoffice
# Ubuntu: sudo apt-get install libreoffice
```

### Running Conversions
```bash
# Basic conversion
python comprehensive_to_markdown_converter.py "/path/to/documents"

# With output directory
python comprehensive_to_markdown_converter.py "/path/to/documents" -o "/path/to/output"

# Advanced options
python comprehensive_to_markdown_converter.py "/path/to/documents" --overwrite --no-recursive --no-ocr

# Quick script (pre-configured path)
./run_comprehensive_converter.sh
```

### Testing
```bash
# Test Docling functionality
python test_docling.py

# Test Excel conversion
python test_excel_conversion.py

# Create test Excel files
python create_test_excel.py
```

## Development Guidelines

### File Processing Logic
- Word files (.doc/.docx) → PDF → Markdown (via LibreOffice + Docling)
- Excel files (.xlsx/.xls) → individual sheet Markdown files (via Pandas)
- Excel sheet names with special characters are sanitized (replaced with underscores)
- Output files named as: `ExcelFile-SheetName.md`

### Logging and Monitoring
- Detailed logging to `.log` files in project root
- Separate logs for different conversion types
- Console output for real-time monitoring

### Virtual Environment Management
- All operations should be done within `venv_docling`
- Dependencies are pinned in `requirements_docling.txt`
- Docling downloads AI models on first run (may take several minutes)

### Error Handling
- LibreOffice dependency checks before Word conversion
- Graceful handling of unsupported file formats
- Memory considerations for large Excel files
- File path sanitization for cross-platform compatibility

## Key Dependencies

- `docling>=2.38.0`: Core document processing
- `pandas>=2.3.0`: Excel data handling
- `openpyxl>=3.1.5`: Excel file parsing
- External: LibreOffice (system dependency)

## Testing Strategy

The project includes test scripts that validate:
- Docling library installation and functionality
- Excel sheet extraction and conversion
- File path handling and output generation