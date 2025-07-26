from docx2pdf import convert
import os

# Replace with your actual directory
docx_dir = "generated_documents"

# Convert all .docx files in the folder
convert(docx_dir)
