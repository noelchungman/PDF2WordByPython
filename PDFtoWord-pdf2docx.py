# pip install pdf2docx
# python PDFtoWord.py

# git init
# git add .
# git commit -m "Initial commit"
# git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPOSITORY.git
# git push -u origin master

# Comment: this library cannot export table in PDF to word

from pdf2docx import Converter

# Paths
input_pdf_path = 'input.pdf'
output_docx_path = 'output.docx'

# Convert PDF to DOCX
cv = Converter(input_pdf_path)
cv.convert(output_docx_path, start=0, end=None)
cv.close()