# PDF Splitter & Text Extractor
This command-line utility splits PDF files based on a code seperator, renames files using the part number, extracts Arabic and English text from PDF files using Tesseract OCR, and writes the pages to an Excel file. The output is used to search for data in files.

## Usage
To use the utility, simply run the following command:
```
python pdf_text_extractor.py <input_path> <output_path>
```
Replace <input_path> with the path to the directory containing the input PDF files and <output_path> with the path to the directory where the output PDF files and Excel file will be saved.

## Requirements
The utility requires the following Python packages:

* os
* warnings
* sys
* time
* click
* pytesseract
* fitz
* PIL
* xlsxwriter
* shutil

It also requires Tesseract OCR. On Windows, you will need to specify the path to Tesseract OCR in extractor.py. The path is usually as follows:
```
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
```

For Linux, comment the line before and type the following in the terminal to install Tesseract OCR:
```
sudo apt install tesseract-ocr
sudo apt install tesseract-ocr-ara
```

## Functionality
The utility works by first validating the input and output paths and then extracting text from each PDF file in the input directory. The text and images are stored in an array, and if a code seperator exists in a page, None is added instead of page image and text. Then, the PDF files are generated based on the None values of the seperators, and text is written to Excel. The text is written to two columns: "File Name" and "Content".

## Limitations
The utility only works with PDF files.
The utility requires using a pre-defined code.
The PDF resolution should be 150 DPIs or more.
The utility has only been tested only on Windows and Ubuntu.
