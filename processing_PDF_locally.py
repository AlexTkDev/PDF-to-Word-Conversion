import fitz  # PyMuPDF
import pytesseract
from PIL import Image
from docx import Document
from docx.shared import Pt, Inches
from tqdm import tqdm
import os


def ocr_image_to_text(image):
    """Преобразование изображения в текст с использованием OCR"""
    return pytesseract.image_to_string(image)


def create_word_document(pdf_path):
    """Создание одного документа Word для всего PDF"""
    if not os.path.exists('output'):
        os.makedirs('output')

    pdf_document = fitz.open(pdf_path)
    num_pages = len(pdf_document)

    doc = Document()

    # Настройка параметров страницы
    section = doc.sections[0]
    section.page_height = Inches(11.69)
    section.page_width = Inches(8.27)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)

    for page_num in tqdm(range(num_pages), desc="Processing pages"):
        page = pdf_document.load_page(page_num)

        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        text = ocr_image_to_text(img)

        for paragraph in text.split('\n'):
            if paragraph.strip():  # Добавляем только непустые строки
                p = doc.add_paragraph(paragraph)
                p.style.font.name = 'Times New Roman'
                p.style.font.size = Pt(14)

                # Форматирование заголовков
                if paragraph.istitle():
                    p.style.font.size = Pt(17)

    # Сохраняем весь документ в один файл
    output_path = 'output/complete_document.docx'
    doc.save(output_path)


def main():
    create_word_document('Open Project 100 images.pdf')


if __name__ == "__main__":
    main()
