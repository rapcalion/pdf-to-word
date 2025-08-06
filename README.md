# Comprehensive PDF to Word Converter

A powerful Python tool that converts PDF files to Word documents while preserving formatting, tables, images, and layout.

## Features

- **Multiple Conversion Methods**:
  - **Hybrid**: Combines multiple libraries for best results (recommended)
  - **pdf2docx**: Fast conversion with good formatting
  - **PyMuPDF**: Handles complex documents well
  - **pdfplumber**: Excellent table extraction

- **Comprehensive Format Support**:
  - Text-based PDFs
  - Scanned PDFs (with OCR)
  - Table-heavy documents
  - Multi-column layouts
  - Images and graphics
  - Forms and annotations
  - Encrypted PDFs (if accessible)

- **Visual Formatting Preservation**:
  - Font styles (bold, italic, underline)
  - Font sizes and families
  - Colors and highlighting
  - Tables with merged cells
  - Image positioning
  - Page breaks
  - Headers and footers

## Installation

1. Create and activate virtual environment:
```bash
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

2. Install dependencies:
```bash
pip install PyMuPDF python-docx pdfplumber pdf2docx pytesseract opencv-python-headless numpy pandas
```

3. For OCR support (scanned PDFs), install Tesseract:
```bash
# macOS
brew install tesseract

# Ubuntu/Debian
sudo apt-get install tesseract-ocr

# Windows
# Download from: https://github.com/UB-Mannheim/tesseract/wiki
```

## Usage

### Simple Usage

```bash
# Convert single PDF (auto-selects best method)
python convert_pdf.py document.pdf

# Interactive mode
python convert_pdf.py
```

### Command Line Options

```bash
# Specify output file
python convert_pdf.py input.pdf -o output.docx

# Choose conversion method
python convert_pdf.py document.pdf -m pdf2docx

# Batch convert entire folder
python convert_pdf.py --batch /path/to/pdf/folder
```

### Python API

```python
from comprehensive_pdf_converter import ComprehensivePDFConverter

# Create converter
converter = ComprehensivePDFConverter()

# Convert PDF
success = converter.convert('input.pdf', 'output.docx', method='hybrid')
```

## Conversion Methods

| Method | Best For | Speed | Quality |
|--------|----------|-------|---------|
| hybrid | Most documents | Medium | Excellent |
| pdf2docx | Standard PDFs | Fast | Good |
| pymupdf | Complex layouts | Medium | Good |
| pdfplumber | Table-heavy docs | Slow | Excellent for tables |

## Troubleshooting

### Common Issues

1. **Tables not detected properly**
   - Try `pdfplumber` method
   - Use `comprehensive_pdf_converter.py` directly for more control

2. **Scanned PDF not converting**
   - Ensure Tesseract is installed
   - The tool auto-detects scanned PDFs and uses OCR

3. **Formatting issues**
   - Try different methods
   - `hybrid` method usually gives best results

4. **Large files**
   - Processing may take time
   - Monitor console for progress

### Tips for Best Results

- Use high-quality PDFs when possible
- For scanned documents, ensure good scan quality
- Try different methods if first attempt isn't satisfactory
- Check console output for warnings about specific issues

## File Structure

```
pdfToWord/
├── convert_pdf.py              # Main CLI tool
├── comprehensive_pdf_converter.py  # Core converter class
├── pdfToWord.py               # Basic converter
├── advanced_converter.py       # Advanced features
└── README.md                  # This file
```

## License

MIT License - Feel free to use and modify as needed.