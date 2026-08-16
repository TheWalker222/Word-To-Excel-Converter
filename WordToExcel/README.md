# Word to Excel Automation

Python tool that extracts structured data and tables from Word documents and exports them to Excel.

## Features

- Batch processing of `.docx` files
- Text extraction
- Table extraction
- Automatic conversion of integers and floats
- Date conversion
- Currency formatting
- Automatic column width
- Excel filters
- Frozen headers
- Error handling
- Logging

## Technologies

- Python
- python-docx
- openpyxl

## Installation

 ```bash```
 - pip install -r requirements.txt
## Usage
 - Put your .docx files into the OfficeFiles folder and run:
 - python word_to_excel.py
 - The converted Excel file will be saved in the output folder.