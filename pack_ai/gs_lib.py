import subprocess
import os

def compress_pdf(input_pdf, output_pdf, pdf_quality="/ebook", verbose=False, gs_command='gs'):
    """
    Сжимает PDF-файл с помощью Ghostscript.

    :param input_pdf: путь к исходному PDF-файлу
    :param output_pdf: путь к выходному (сжатому) PDF-файлу
    :param pdf_quality: уровень сжатия:
        - "/screen"  — низкое качество, минимальный размер
        - "/ebook"   — среднее качество
        - "/printer" — высокое качество (меньше сжатия)
        - "/prepress" — максимально возможное качество
    :param verbose: выводить ли подробное сообщение
    :raises: RuntimeError, если файл не найден или ошибка в Ghostscript
    """
    if not os.path.exists(input_pdf):
        raise FileNotFoundError(f"Файл не найден: {input_pdf}")

    # Параметры сжатия
    args = [
        gs_command, # gs gswin64c.exe
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.4",
        f"-dPDFSETTINGS={pdf_quality}",
        "-dNOPAUSE",
        "-dQUIET",
        "-dBATCH",
        f"-sOutputFile={output_pdf}",
        input_pdf
    ]

    if verbose:
        print(f"Сжатие PDF: {input_pdf} → {output_pdf} (качество: {pdf_quality})")

    try:
        result = subprocess.run(args, check=True, capture_output=True)
        if verbose:
            print(f"Успешно сжато. Размер исходного: {os.path.getsize(input_pdf)} байт")
            print(f"Размер сжатого: {os.path.getsize(output_pdf)} байт")
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"Ошибка Ghostscript: {e.stderr.decode()}")
    except FileNotFoundError:
        raise RuntimeError("Ghostscript не установлен или не найден в PATH. Скачайте с https://www.ghostscript.com/download/gsdnld.html")


# === Пример использования ===
if __name__ == "__main__":
    compress_pdf(
        input_pdf="результат_20250820-163541.pdf",
        output_pdf="результат_20250820-163541-compressed.pdf",
        pdf_quality="/ebook",  # можно попробовать "/screen" для максимального сжатия
        verbose=True
    )