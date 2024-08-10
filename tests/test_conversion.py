import os
import tempfile
from processing_PDF_locally import ocr_image_to_text, create_word_document
from PIL import Image
import fitz  # PyMuPDF


def test_ocr_image_to_text():
    # Создаем временный файл изображения
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as temp_image_file:
        temp_image_path = temp_image_file.name

        # Создаем простое изображение для теста
        image = Image.new('RGB', (100, 100), color='white')
        image.save(temp_image_path)

        # Тестируем функцию ocr_image_to_text с временным изображением
        result = ocr_image_to_text(temp_image_path)
        assert isinstance(result, str)

        # Удаляем временный файл
        os.remove(temp_image_path)


def test_create_word_document():
    # Создаем временный файл PDF
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_pdf_file:
        temp_pdf_path = temp_pdf_file.name

        # Создаем простой PDF для теста
        pdf_document = fitz.open()
        pdf_document.new_page()
        pdf_document.save(temp_pdf_path)
        pdf_document.close()

        # Тестируем функцию create_word_document с временным PDF
        output_path = 'output/complete_document.docx'
        create_word_document(temp_pdf_path)

        # Проверяем, что файл был создан
        assert os.path.exists(output_path)

        # Удаляем временные файлы
        os.remove(temp_pdf_path)
        os.remove(output_path)
