import os
from docx2pdf import convert
from PyPDF2 import PdfMerger
import tempfile
import shutil
from pack.datelib import *
#from rich import print

from pack_ai.gs_lib import compress_pdf
from pack.docx_to_pdf_linux_lib import *

# Пример использования
if __name__ == "__main__":
    #folder = "путь/к/вашей/папке"  # ← замените на реальный путь
    folder = './in'
    #output = "результат.pdf" 
    output =  genfname(pref='результат_',postf='.pdf')
    merge_docx_to_pdf(folder, output)
    output_compressed = output.replace('.pdf', '_compressed_.pdf')
    #compress_pdf(input_pdf=output, output_pdf=output_compressed)
    #print(f"Объединённый PDF сжатый: {output_compressed}")
    input("press any key...")
    