from docx2pdf import convert
import subprocess
import os
import tempfile

def convert_docx_to_pdf_compressed(input_docx, output_pdf, 
                                   pdf_quality="/ebook",
                                   cleanup_temp=True):
    """
    Конвертирует .docx файл в PDF и сжимает его с помощью Ghostscript.
    
    :param input_docx: путь к входному .docx файлу
    :param output_pdf: путь к выходному (сжатому) PDF файлу
    :param pdf_quality: уровень сжатия:
        - "/screen"  — низкое качество, самый маленький размер
        - "/ebook"   — среднее качество
        - "/printer" — высокое качество
        - "/prepress" — очень высокое (не для сжатия)
    :param cleanup_temp: удалить ли временный PDF после сжатия
    """
    # Создаём временный файл для промежуточного PDF
    temp_pdf = tempfile.mktemp(suffix=".pdf")

    try:
        # Шаг 1: Конвертация docx -> PDF
        print("Конвертация DOCX в PDF...")
        convert(input_docx, temp_pdf)

        # Шаг 2: Сжатие PDF через Ghostscript
        print("Сжатие PDF...")
        subprocess.run([
            "gswin64c.exe",
            "-sDEVICE=pdfwrite",
            "-dCompatibilityLevel=1.4",
            f"-dPDFSETTINGS={pdf_quality}",
            "-dNOPAUSE",
            "-dQUIET",
            "-dBATCH",
            f"-sOutputFile={output_pdf}",
            temp_pdf], check=True)

        print(f"Готово! Сжатый PDF сохранён: {output_pdf}")

    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"Ошибка при сжатии PDF через Ghostscript: {e}")
    except Exception as e:
        raise RuntimeError(f"Ошибка при конвертации DOCX в PDF: {e}")
    finally:
        # Удаление временного файла
        if cleanup_temp and os.path.exists(temp_pdf):
            os.remove(temp_pdf)


# === Пример использования ===
if __name__ == "__main__":
    convert_docx_to_pdf_compressed(
        input_docx="document.docx",
        output_pdf="document_compressed.pdf",
        pdf_quality="/ebook"  # баланс между качеством и размером
    )