## PDF-to-Word-Conversion

### Description

The **PDF-to-Word-Conversion** project is designed to convert pages of a PDF document into separate
pages of a Microsoft
Word document while adhering to specific technical requirements. The project employs two main
processing strategies:
using the Google Cloud Vision API for text extraction and local processing with PyMuPDF and
Tesseract.

### Tech Stack

- **Google Cloud Vision API**: For extracting text from images.
- **PyMuPDF**: For working with PDF documents.
- **Tesseract OCR**: For recognizing text in images.
- **pdf2image**: For converting PDF pages into images.
- **Pillow**: For image processing.
- **python-docx**: For creating and editing Word documents.
- **tqdm**: For displaying progress.
- **Flake8**: For linting.
- **Pytest**: For tests.

### Installation

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
        - Add the path to Tesseract in the PATH environment variable (e.g.,
          `C:\Program Files\Tesseract-OCR`).

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
         pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
            # for Windows
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

### Running tests

- To run tests, use the command:

```bash
 pytest tests/
```

### Linting

- To check your code for standards compliance, use Flake8. Run the command:

```bash
flake8 PDF_processing_with_Google_Cloud_Vision_API.py processing_PDF_locally.py --show-source --statistics
flake8 PDF_processing_with_Google_Cloud_Vision_API.py processing_PDF_locally.py --count --exit-zero --max-complexity=10 --max-line-length=127 --statistics
```

### Notes

- In the script `PDF_processing_with_Google_Cloud_Vision_API.py`, replace
  `'path_to_your_google_credentials.json'` with
  the path to your Google Cloud credentials file.
- In the script `processing_PDF_locally.py`, replace `'Open Project 100 images.pdf'` with the path
  to your PDF document.

### Project Structure

- **PDF_processing_with_Google_Cloud_Vision_API.py**: Script for extracting text from images using
  Google Cloud Vision
  API and creating a Word document.
- **processing_PDF_locally.py**: Script for extracting text from images using Tesseract and
  creating a Word document.

### Contributing

- If you want to make changes to the project, fork the repository, make your changes, and create a
  Pull Request.

### License

This project is licensed under the [MIT License](https://opensource.org/licenses/MIT).
