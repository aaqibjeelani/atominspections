# Word Document Merger (Python Version)

A Python application that lets you merge multiple Word documents while applying automatic sequential numbering to table cells in the first column (REF column).

## Features

- Merge multiple Word documents (.docx files)
- Preserve formatting from template document (margins, headers, footers)
- Add sequential numbering to the first column (REF column) in all tables
- Skip header rows (1st row, 5th row, 9th row, etc.)
- User-friendly GUI for easy operation
- Command-line interface for batch processing

## Requirements

- Python 3.6 or higher
- python-docx library

## Installation

1. Make sure you have Python installed on your system
2. Install required dependencies:

```bash
pip install -r requirements.txt
```

## Usage

### GUI Mode

To use the graphical interface:

```bash
python merge_word_docs.py
```

1. Select your template document (provides formatting, headers, footers)
2. Add documents to merge
3. Specify the output file location
4. Click "Merge Documents"

### Command Line Mode

For command-line usage:

```bash
python merge_word_docs.py <template.docx> <output.docx> <doc1.docx> [doc2.docx] [...]
```

For example:
```bash
python merge_word_docs.py sample/template.docx merged.docx document1.docx document2.docx
```

## How It Works

1. Takes formatting and layout from the template document
2. Merges content from all specified documents
3. Processes all tables in the merged document:
   - Skips header rows (1st row, 5th row, 9th row, etc.)
   - Adds sequential numbers in the first column (REF column)
   - Preserves table formatting

## Troubleshooting

If you encounter any issues:

1. Make sure all documents are valid .docx files
2. Ensure the template has proper formatting
3. Check console output for detailed error messages 