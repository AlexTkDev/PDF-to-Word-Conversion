import io
import os
from google.cloud import vision
from google.cloud.vision_v1 import types
from pdf2image import convert_from_path
from docx import Document
from docx.shared import Pt, Inches
from tqdm import tqdm

# Путь к файлу JSON с учетными данными Google Cloud
os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = 'path_to_your_google_credentials.json'


def detect_text_google_cloud(image):
    client = vision.ImageAnnotatorClient()
    content = image.tobytes()
    image = types.Image(content=content)
    response = client.text_detection(image=image)
    texts = response.text_annotations
    detected_text = ""
    for text in texts:
        detected_text += text.description + "\n"
    return detected_text


def preprocess_image(image):
    image = image.convert("L")  # Преобразование в черно-белый формат
    threshold = 128
    image = image.point(lambda p: p > threshold and 255)  # Бинаризация
    return image


def create_word_document(pdf_path):
    if not os.path.exists('output'):
        os.makedirs('output')

    # Конвертация PDF в изображения
    images = convert_from_path(pdf_path)

    # Создаем один документ Word
    doc = Document()

    # Настраиваем параметры страницы
    section = doc.sections[0]
    section.page_height = Inches(11.69)
    section.page_width = Inches(8.27)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)

    for page_num in tqdm(range(len(images)), desc="Processing pages"):
        image = images[page_num]
        image = preprocess_image(image)
        text = detect_text_google_cloud(image)

        for paragraph in text.split('\n'):
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
    pdf_path = 'your_pdf_path.pdf'
    create_word_document(pdf_path)


if __name__ == "__main__":
    main()
