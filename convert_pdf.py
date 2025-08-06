#!/usr/bin/env python3
"""
Simple PDF to Word Converter
Easy-to-use interface for comprehensive PDF conversion
"""

import os
import sys
from comprehensive_pdf_converter import ComprehensivePDFConverter
import argparse


def convert_pdf_to_word(pdf_path, output_path=None, method='hybrid'):
    """
    Convert PDF to Word with best quality
    """
    # Generate output path if not provided
    if not output_path:
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_path = f"{base_name}_converted.docx"
    
    # Create converter
    converter = ComprehensivePDFConverter()
    
    print(f"Converting: {pdf_path}")
    print(f"Method: {method}")
    print(f"Output: {output_path}")
    print("-" * 50)
    
    # Convert
    success = converter.convert(pdf_path, output_path, method)
    
    if success:
        print(f"\n✓ Success! Saved to: {output_path}")
        file_size = os.path.getsize(output_path) / 1024 / 1024  # MB
        print(f"File size: {file_size:.2f} MB")
    else:
        print("\n✗ Conversion failed!")
        print("Tips:")
        print("- Try method 'pdf2docx' for standard PDFs")
        print("- Try method 'pdfplumber' for table-heavy PDFs")
        print("- Check if PDF is corrupted or encrypted")
    
    return success


def batch_convert(folder_path, method='hybrid'):
    """
    Convert all PDFs in a folder
    """
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        print("No PDF files found in the folder!")
        return
    
    print(f"Found {len(pdf_files)} PDF files")
    print("-" * 50)
    
    success_count = 0
    
    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        print(f"\nProcessing: {pdf_file}")
        
        if convert_pdf_to_word(pdf_path, method=method):
            success_count += 1
    
    print("\n" + "=" * 50)
    print(f"Batch conversion complete: {success_count}/{len(pdf_files)} successful")


def main():
    parser = argparse.ArgumentParser(
        description='Convert PDF to Word while preserving formatting',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Convert single file (auto-detect best method)
  python convert_pdf.py document.pdf
  
  # Convert with specific method
  python convert_pdf.py document.pdf -m pdf2docx
  
  # Convert to specific output file
  python convert_pdf.py input.pdf -o output.docx
  
  # Batch convert folder
  python convert_pdf.py --batch /path/to/folder
  
Methods:
  hybrid     - Best quality, combines multiple methods (default)
  pdf2docx   - Fast, good for standard PDFs
  pymupdf    - Good for complex documents
  pdfplumber - Best for table-heavy PDFs
        """
    )
    
    parser.add_argument('input', nargs='?', help='PDF file path or folder (with --batch)')
    parser.add_argument('-o', '--output', help='Output DOCX file path')
    parser.add_argument('-m', '--method', 
                       choices=['hybrid', 'pdf2docx', 'pymupdf', 'pdfplumber'],
                       default='hybrid',
                       help='Conversion method (default: hybrid)')
    parser.add_argument('--batch', action='store_true', 
                       help='Batch convert all PDFs in folder')
    
    args = parser.parse_args()
    
    # Interactive mode if no arguments
    if not args.input:
        print("PDF to Word Converter")
        print("=" * 50)
        
        if input("Batch conversion? (y/n): ").lower() == 'y':
            folder = input("Enter folder path: ").strip()
            if os.path.isdir(folder):
                batch_convert(folder)
            else:
                print("Invalid folder path!")
        else:
            pdf_path = input("Enter PDF file path: ").strip()
            if os.path.isfile(pdf_path):
                convert_pdf_to_word(pdf_path)
            else:
                print("File not found!")
        return
    
    # Command line mode
    if args.batch:
        if os.path.isdir(args.input):
            batch_convert(args.input, args.method)
        else:
            print("Error: --batch requires a folder path")
    else:
        if os.path.isfile(args.input):
            convert_pdf_to_word(args.input, args.output, args.method)
        else:
            print(f"Error: File not found: {args.input}")


if __name__ == "__main__":
    main()