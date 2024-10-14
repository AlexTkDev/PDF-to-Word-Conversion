import os
from google.cloud import vision
from google.cloud.vision_v1 import types
from pdf2image import convert_from_path
from docx import Document
from docx.shared import Pt, Inches
from tqdm import tqdm

# Установка пути к файлу JSON с учетными данными Google Cloud
os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = 'path_to_your_google_credentials.json'


class PDFToWordConverter:
    def __init__(self, pdf_path):
        self.pdf_path = pdf_path
        self.client = vision.ImageAnnotatorClient()

    def detect_text_google_cloud(self, image):
        """Обнаружение текста в изображении с помощью Google Cloud Vision API."""
        content = image.tobytes()
        vision_image = types.Image(content=content)

        try:
            response = self.client.text_detection(image=vision_image)
            texts = response.text_annotations
            detected_text = "\n".join(text.description for text in texts)
            return detected_text
        except Exception as e:
            print(f"Error during text detection: {e}")
            return ""

    def preprocess_image(self, image):
        """Преобразование изображения в черно-белый формат и бинаризация."""
        image = image.convert("L")  # Преобразование в черно-белый формат
        threshold = 128
        image = image.point(lambda p: p > threshold and 255)  # Бинаризация
        return image

    def create_word_document(self):
        """Создание документа Word из PDF файла."""
        if not os.path.exists('output'):
            os.makedirs('output')

        # Конвертация PDF в изображения
        images = convert_from_path(self.pdf_path)

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
            processed_image = self.preprocess_image(image)
            text = self.detect_text_google_cloud(processed_image)

            for paragraph in text.split('\n'):
                if paragraph.strip():  # Проверка на пустые строки
                    p = doc.add_paragraph(paragraph)
                    p.style.font.name = 'Times New Roman'
                    p.style.font.size = Pt(14)

                    # Форматирование заголовков
                    if paragraph.istitle():
                        p.style.font.size = Pt(17)

        # Сохраняем весь документ в один файл
        output_path = 'output/complete_document.docx'
        doc.save(output_path)
        print(f"Document saved to {output_path}")


def main():
    # Укажите путь к PDF файлу
    pdf_path = 'your_pdf_path.pdf'
    converter = PDFToWordConverter(pdf_path)
    converter.create_word_document()


if __name__ == "__main__":
    main()
