# SpectrumMark
ENGLISH[README.MD](https://github.com/gaisuwen/SpectrumMark/edit/main/README.md) | 中文[README_CN.MD](https://github.com/gaisuwen/SpectrumMark/edit/main/README_CN.md)

SpectrumMark is a tool designed for batch-adding captions and automatically generating bookmarks for spectra in pharmaceutical registration submissions.  
This tool can batch-insert specified text into each page of a PDF spectrum file at a designated position and automatically generate a bookmark directory for each page, making it easier to organize and review submission materials.

## Features

- Batch add captions/annotations to each page of a PDF spectrum file
- Support for custom insertion position, font size, color, and alignment
- Automatically generate PDF bookmarks based on the inserted text
- Easy-to-use graphical user interface

## Typical Use Cases

- Organizing and annotating spectrum files for pharmaceutical registration submissions
- Batch adding annotations and directories to spectrum or test report PDF files

## Usage

1. Install dependencies  
   You need to install [PyMuPDF](https://pymupdf.readthedocs.io/) (fitz) first:
   ```
   pip install pymupdf
   ```

2. Run the program  
   ```
   python SpectrumMark.py
   ```

3. Operation steps  
   - Select the input PDF file (spectrum)
   - Select the text file to insert (each line corresponds to a caption/annotation for one page)
   - Select the output PDF file path
   - Set the insertion coordinates, font size, color, and alignment parameters
   - Click "Start", and wait for the process to complete

## Parameter Description

- **Input PDF file**: The spectrum PDF to be batch-captioned
- **Text file to insert**: Each line corresponds to the caption/annotation to be inserted on each page (UTF-8 encoding)
- **Output PDF file**: The newly generated PDF file
- **X/Y coordinates**: The starting position for inserting text (unit: pixels, origin at the bottom left)
- **Font size**: Font size for the inserted text
- **Text color**: RGB format, e.g., `0,0,0` (black), supports 0-1 or 0-255 range
- **Alignment**: left/center/right

## Main File

- [SpectrumMark.py](SpectrumMark.py): Main program, contains all features and UI

## Notes

- The number of lines in the text file should match the number of PDF pages; extra lines will be ignored
- Supports mixed Chinese and English text
- Ensure the `china-s` font is available (built-in with PyMuPDF)

---

For questions or suggestions, feel free to contact us.