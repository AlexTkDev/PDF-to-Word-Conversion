# PDF-to-Word-Conversion

## Description

The **PDF-to-Word-Conversion** project is designed to convert pages of a PDF document into separate pages of a Microsoft
Word document while adhering to specific technical requirements. The project employs two main processing strategies:
using the Google Cloud Vision API for text extraction and local processing with PyMuPDF and Tesseract.

## Tech Stack

- **Google Cloud Vision API**: For extracting text from images.
- **PyMuPDF**: For working with PDF documents.
- **Tesseract OCR**: For recognizing text in images.
- **pdf2image**: For converting PDF pages into images.
- **Pillow**: For image processing.
- **python-docx**: For creating and editing Word documents.
- **tqdm**: For displaying progress.

## Installation

1. **Clone the repository:**

   ```bash
   git clone https://github.com/AlexTkDev/PDF-to-Word-Conversion.git
   cd PDF-to-Word-Conversion
   ```

2. **Create and activate a virtual environment:**

   ```bash
   python -m venv venv
   source venv/bin/activate  # For Windows use `venv\Scripts\activate`
   ```

3. **Install dependencies:**

   ```bash
   pip install -r requirements.txt
   ```

   **Note:** Ensure that **Google Cloud SDK** and **Tesseract OCR** are installed and configured.

4. **Set up Google Cloud Vision API:**

    - Obtain a Google Cloud credentials file (JSON format).
    - Set the path to this file in the environment variable `GOOGLE_APPLICATION_CREDENTIALS`.

5. **Setting up Tesseract OCR:**

    - **For Windows:**
        - Download the Tesseract OCR installer from
          the [official repository](https://github.com/tesseract-ocr/tesseract)
          or [source](https://github.com/UB-Mannheim/tesseract/wiki).
        - Run the installer and follow the instructions.
        - Add the path to Tesseract in the PATH environment variable (e.g., `C:\Program Files\Tesseract-OCR`).

    - **For macOS:**
        - Use Homebrew to install Tesseract. Open a terminal and run:

          ```bash
          brew install tesseract
          ```

    - **For Linux:**
        - Install Tesseract using a package manager. For example, on Ubuntu:

          ```bash
          sudo apt-get update
          sudo apt-get install tesseract-ocr
          ```

    - **Installing Language Packages:**
        - To add additional language packages, follow the instructions for your operating system.

    - **Using in Python:**
        - Install the `pytesseract` library:

          ```bash
          pip install pytesseract
          ```

        - Ensure that `pytesseract` knows where Tesseract is installed:

          ```python
          import pytesseract
      
          # Specify the path to Tesseract if it's not in PATH
          pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  # for Windows
          ```

6. **Run the scripts:**

    - To process PDF using Google Cloud Vision API:

      ```bash
      python PDF_processing_with_Google_Cloud_Vision_API.py
      ```

    - To process PDF locally using PyMuPDF and Tesseract:

      ```bash
      python processing_PDF_locally.py
      ```

## Running tests

- To run tests, use the command:

```bash
        Copy code
        pytest tests/
```

## Linting

- To check your code for standards compliance, use Flake8. Run the command:

```bash
Copy code
flake8 PDF_processing_with_Google_Cloud_Vision_API.py processing_PDF_locally.py --show-source --statistics
flake8 PDF_processing_with_Google_Cloud_Vision_API.py processing_PDF_locally.py --count --exit-zero --max-complexity=10 --max-line-length=127 --statistics
```

## Notes

- In the script `PDF_processing_with_Google_Cloud_Vision_API.py`, replace `'path_to_your_google_credentials.json'` with
  the path to your Google Cloud credentials file.
- In the script `processing_PDF_locally.py`, replace `'Open Project 100 images.pdf'` with the path to your PDF document.

## Project Structure

- **PDF_processing_with_Google_Cloud_Vision_API.py**: Script for extracting text from images using Google Cloud Vision
  API and creating a Word document.
- **processing_PDF_locally.py**: Script for extracting text from images using Tesseract and creating a Word document.

## Contributing

- If you want to make changes to the project, fork the repository, make your changes, and create a Pull Request.

## License

This project is licensed under the [MIT License](https://opensource.org/licenses/MIT).

***

# PDF-to-Word-Conversion

## Описание

Проект **PDF-to-Word-Conversion** предназначен для конвертации страниц PDF-документа в отдельные страницы документа
Microsoft Word с соблюдением определенных технических требований. Проект использует две основные стратегии обработки: с
использованием Google Cloud Vision API для извлечения текста и локальную обработку с помощью PyMuPDF и Tesseract.

## Стек технологий

- **Google Cloud Vision API**: Для извлечения текста из изображений.
- **PyMuPDF**: Для работы с PDF-документами.
- **Tesseract OCR**: Для распознавания текста на изображениях.
- **pdf2image**: Для конвертации страниц PDF в изображения.
- **Pillow**: Для обработки изображений.
- **python-docx**: Для создания и редактирования документов Word.
- **tqdm**: Для отображения прогресса выполнения.

## Установка

1. **Клонируйте репозиторий:**

   ```bash
   git clone https://github.com/AlexTkDev/PDF-to-Word-Conversion.git
   cd PDF-to-Word-Conversion
   ```

2. **Создайте и активируйте виртуальное окружение:**

   ```bash
   python -m venv venv
   source venv/bin/activate  # Для Windows используйте `venv\Scripts\activate`
   ```

3. **Установите зависимости:**

   ```bash
   pip install -r requirements.txt
   ```

   **Примечание:** Убедитесь, что у вас установлены и настроены **Google Cloud SDK** и **Tesseract OCR**.

4. **Настройте Google Cloud Vision API:**

    - Получите файл учетных данных Google Cloud (формат JSON).
    - Задайте путь к этому файлу в переменной окружения `GOOGLE_APPLICATION_CREDENTIALS`.

5. **Настройка Tesseract OCR:**

    - **Для Windows:**
        - Скачайте установочный файл Tesseract OCR
          с [официального репозитория](https://github.com/tesseract-ocr/tesseract)
          или [источника](https://github.com/UB-Mannheim/tesseract/wiki).
        - Запустите установочный файл и следуйте инструкциям.
        - Добавьте путь к Tesseract в переменную среды PATH (например, `C:\Program Files\Tesseract-OCR`).

    - **Для macOS:**
        - Используйте Homebrew для установки Tesseract. Откройте терминал и выполните:

          ```bash
          brew install tesseract
          ```

    - **Для Linux:**
        - Установите Tesseract через пакетный менеджер. Например, для Ubuntu:

          ```bash
          sudo apt-get update
          sudo apt-get install tesseract-ocr
          ```

    - **Установка языковых пакетов:**
        - Для добавления дополнительных языковых пакетов следуйте инструкциям для вашей операционной системы.

    - **Использование в Python:**
        - Установите библиотеку `pytesseract`:

          ```bash
          pip install pytesseract
          ```

        - Убедитесь, что `pytesseract` знает, где установлен Tesseract:

          ```python
          import pytesseract
   
          # Укажите путь к Tesseract, если он не находится в PATH
          pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  # для Windows
          ```

6. **Запустите скрипты:**

    - Для обработки PDF с помощью Google Cloud Vision API:

      ```bash
      python PDF_processing_with_Google_Cloud_Vision_API.py
      ```

    - Для локальной обработки PDF с помощью PyMuPDF и Tesseract:

      ```bash
      python processing_PDF_locally.py
      ```

## Запуск тестов

- Для запуска тестов используйте команду:

```bash
Копировать код
pytest tests/
```

## Linting

- Для проверки кода на соответствие стандартам используйте Flake8. Запустите команду:

```bash
Копировать код
flake8 PDF_processing_with_Google_Cloud_Vision_API.py processing_PDF_locally.py --show-source --statistics
flake8 PDF_processing_with_Google_Cloud_Vision_API.py processing_PDF_locally.py --count --exit-zero --max-complexity=10 --max-line-length=127 --statistics
```

## Примечания

- В скрипте `PDF_processing_with_Google_Cloud_Vision_API.py` необходимо
  заменить `'path_to_your_google_credentials.json'` на путь к вашему файлу с учетными данными Google Cloud.
- В скрипте `processing_PDF_locally.py` замените `'Open Project 100 images.pdf'` на путь к вашему PDF-документу.

## Структура проекта

- **PDF_processing_with_Google_Cloud_Vision_API.py**: Скрипт для извлечения текста из изображений с помощью Google Cloud
  Vision API и создания документа Word.
- **processing_PDF_locally.py**: Скрипт для извлечения текста из изображений с помощью Tesseract и создания документа
  Word.

## Контрибьюция

- Если вы хотите внести изменения в проект, создайте форк репозитория, сделайте ваши изменения и создайте Pull Request.

## Лицензия

Этот проект лицензируется под [MIT License](https://opensource.org/licenses/MIT).