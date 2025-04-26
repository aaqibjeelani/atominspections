# Word Document Merger

A powerful tool for merging multiple Word documents while preserving formatting and adding sequential numbering to tables.

**Developed by Aaqib Jeelani** - [instagram.com/aaqibjeelani](https://instagram.com/aaqibjeelani)

## Features

- Merge multiple Word (.docx) documents into a single file
- Preserve formatting from a template
- Automatically add sequential REF numbers to the first column of tables
- Skip header rows (first row and every 4th row thereafter)
- Preserve images in tables at their original size
- Standalone executable - no installation required

## Using the Application

1. **Download and unzip** the application package
2. **Double-click** on `WordDocumentMerger.exe` to start the application
3. The application uses a template from the `sample` folder - ensure this folder is present alongside the executable
4. **Click "Add Documents"** to select the Word documents you want to merge
5. **Click "Browse..."** under "Output Document" to specify where to save the merged file
6. **Click "Merge Documents"** to process and create the merged document

## Notes for Advanced Users

- The template is located at `sample/template.docx` - you can replace this with your own template
- Tables should have a first column for REF numbers
- Images in tables are preserved at their original size
- Header rows (1st, 5th, 9th, etc.) are skipped for numbering

## Building from Source

If you want to build the executable yourself:

1. Install Python 3.6 or later
2. Run `pip install -r requirements.txt` to install dependencies
3. Run `python build_exe.py` to create the executable

## Troubleshooting

- If the application doesn't start, ensure you have extracted all files from the zip archive
- Make sure the `sample` directory with the template is alongside the executable
- For any issues, contact the developer

---

© 2024 Aaqib Jeelani. All rights reserved. 