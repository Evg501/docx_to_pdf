import os
from docx2pdf import convert
from PyPDF2 import PdfMerger
import tempfile
import shutil


def merge_docx_to_pdf(folder_path, output_pdf):
    # Проверяем, что папка существует
    if not os.path.isdir(folder_path):
        raise ValueError(f"Папка не найдена: {folder_path}")

    # Получаем список .docx файлов в папке
    docx_files = sorted([
        os.path.join(folder_path, f) for f in os.listdir(folder_path)
        if f.lower().endswith('.docx') and os.path.isfile(os.path.join(folder_path, f))
    ])

    if not docx_files:
        raise ValueError("В папке нет файлов .docx")

    # Создаём временную папку для PDF-файлов
    with tempfile.TemporaryDirectory() as temp_dir:
        pdf_files = []

        # Конвертируем каждый .docx в PDF
        for docx_file in docx_files:
            pdf_filename = os.path.basename(docx_file).rsplit('.', 1)[0] + '.pdf'
            pdf_path = os.path.join(temp_dir, pdf_filename)
            print(f"Конвертируем: {docx_file} → {pdf_path}")
            convert(docx_file, pdf_path)
            pdf_files.append(pdf_path)

        # Объединяем все PDF
        merger = PdfMerger()
        for pdf in pdf_files:
            merger.append(pdf)

        # Сохраняем итоговый PDF
        merger.write(output_pdf)
        merger.close()
        print(f"Объединённый PDF сохранён: {output_pdf}")