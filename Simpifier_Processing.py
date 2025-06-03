import PyPDF2
from docx import Document
import re
from collections import Counter
import pdfplumber
from pdf2image import convert_from_path

import pytesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


def extract_text_from_pdf(pdf_path, use_generic=False, use_pdfplumber=False, use_OCR=False):
    suffixes = ""
    """Extracts text from PDF and removes repeated headers, footers, and page numbers."""
    # In case user unchecks all boxes
    if not use_generic and not use_pdfplumber and not use_OCR:
        use_generic = True

    if use_generic:
        try:
            suffixes += "generic"
            with open(pdf_path, 'rb') as pdf_file_obj:
                pdf_reader = PyPDF2.PdfReader(pdf_file_obj)
                page_texts = []
                headers = []
                footers = []

                # Step 1: Collect headers and footers
                for page in pdf_reader.pages:
                    text = page.extract_text()
                    if not text:
                        continue
                    lines = text.split('\n')
                    if len(lines) >= 2:
                        headers.append(lines[0].strip())
                        footers.append(lines[-1].strip())
                    page_texts.append(lines)

                # Step 2: Find the most common header and footer
                common_header = Counter(headers).most_common(1)[0][0] if headers else ""
                common_footer = Counter(footers).most_common(1)[0][0] if footers else ""

                # Step 3: Reconstruct body text without headers/footers/page numbers
                full_text = ""
                for lines in page_texts:
                    # Remove header and footer if matched
                    if lines and lines[0].strip() == common_header:
                        lines = lines[1:]
                    if lines and lines[-1].strip() == common_footer:
                        lines = lines[:-1]

                    # Remove standalone page numbers
                    lines = [line for line in lines if not line.strip().isdigit()]

                    full_text += ' '.join(lines) + ' '

                return full_text.strip(), suffixes

        except FileNotFoundError:
            return "Error: PDF file not found.", suffixes
        except Exception as e:
            return f"Error reading PDF: {e}", suffixes

    if use_pdfplumber:
        """Extracts text using pdfplumber, better layout interpretation."""
        try:
            suffixes += "plumber"
            full_text = ""
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        lines = text.split('\n')
                        lines = [line for line in lines if not line.strip().isdigit()]  # remove page numbers
                        full_text += ' '.join(lines) + ' '
            return full_text.strip(), suffixes
        except Exception as e:
            return f"Error reading PDF with pdfplumber: {e}", suffixes

    if use_OCR:
        """Extracts text from PDF using OCR on image-rendered pages."""
        try:
            suffixes += "ocr"
            images = convert_from_path(pdf_path)
            full_text = ""
            for img in images:
                text = pytesseract.image_to_string(img)
                full_text += text + " "
            return full_text.strip(), suffixes
        except Exception as e:
            return f"Error during OCR extraction: {e}", suffixes

def clean_text(text, newline, hyphen, doublespace, paragraph, suffixes):
    """Removes extraneous newlines and other potential artifacts."""
    # Replace multiple newlines with a single space
    # This helps join lines that were broken for PDF formatting

    if newline:
        text = re.sub(r'\n+', ' ', text)
        suffixes += "1"

    # Optional: Remove hyphenation at the end of lines
    if hyphen:
        text = re.sub(r'-\s+', '', text) # Uncomment if you have many hyphenated words
        suffixes += "2"

    # Optional: Replace multiple spaces with a single space
    if doublespace:
        text = re.sub(r'\s{2,}', ' ', text)
        suffixes += "3"

    # Heuristic: Treat sentence endings followed by uppercase as paragraph breaks
    if paragraph:
        text = re.sub(r'(?<=[.!?])\s+(?=[A-Z])', '\n\n', text)
        suffixes += "4"

    return text.strip(), suffixes

def save_text_to_word(text, word_path):
    """Saves the given text to a Word document."""
    try:
        document = Document()
        document.add_paragraph(text)
        document.save(word_path)
        return f"Successfully saved to {word_path}"
    except Exception as e:
        return f"Error saving Word document: {e}"
