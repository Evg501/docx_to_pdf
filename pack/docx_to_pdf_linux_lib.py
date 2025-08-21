import os
import subprocess
import tempfile
from PyPDF2 import PdfMerger

def merge_docx_to_pdf(folder_path, output_pdf):
    # Проверка папки
    if not os.path.isdir(folder_path):
        raise ValueError(f"Папка не найдена: {folder_path}")

    # Получаем .docx файлы (в алфавитном порядке)
    docx_files = sorted([
        os.path.join(folder_path, f) for f in os.listdir(folder_path)
        if f.lower().endswith('.docx') and os.path.isfile(os.path.join(folder_path, f))
    ])

    if not docx_files:
        raise ValueError("В папке нет файлов .docx")

    # Временная директория для PDF
    with tempfile.TemporaryDirectory() as temp_dir:
        pdf_files = []

        # Конвертация каждого .docx в PDF через LibreOffice
        for docx_file in docx_files:
            pdf_path = os.path.join(temp_dir, os.path.basename(docx_file).rsplit('.', 1)[0] + '.pdf')
            print(f"Конвертируем: {docx_file} → {pdf_path}")
            result = subprocess.run([
                'libreoffice',
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', temp_dir,
                docx_file
            ], capture_output=True)

            if result.returncode != 0:
                print(f"Ошибка при конвертации {docx_file}: {result.stderr.decode('utf-8')}")
                continue

            # LibreOffice сохраняет PDF как <имя_файла>.pdf
            generated_pdf = os.path.join(temp_dir, os.path.splitext(os.path.basename(docx_file))[0] + '.pdf')
            if os.path.exists(generated_pdf):
                pdf_files.append(generated_pdf)
            else:
                print(f"Не удалось создать PDF для {docx_file}")

        if not pdf_files:
            raise ValueError("Не удалось конвертировать ни один файл в PDF")

        # Объединяем все PDF
        merger = PdfMerger()
        for pdf in pdf_files:
            merger.append(pdf)

        merger.write(output_pdf)
        merger.close()
        print(f"✅ Объединённый PDF сохранён: {output_pdf}")
        