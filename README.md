# DOCX_Generator
creates docx files based on the provided data from CSV file

## How to run
just run `file_creator.py` to generate docx files based on the provided csv and template file.
then run `pdf_converter.py` to generate PDF files from generated docx files. (pdf converter only works on *windows*)

## Creating an exe file
```bash
pyinstaller --onefile <file>.py
```