# pip install pdf2docx
# python PDFtoWord.py

# Comment: this library cannot export table in PDF to word, better than PyMuPDF and pdfplumber
# pdf2docx > PyMuPDF > pdfplumber

from pdf2docx import Converter

# Paths
input_pdf_path = 'input.pdf'
output_docx_path = 'output.docx'

# Convert PDF to DOCX
cv = Converter(input_pdf_path)
cv.convert(output_docx_path, start=0, end=None)
cv.close()