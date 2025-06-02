import PyPDF2
from docx import Document
import re

def extract_text_from_pdf(pdf_path):
    """Extracts text from a PDF file."""
    text = ""
    try:
        with open(pdf_path, 'rb') as pdf_file_obj:
            pdf_reader = PyPDF2.PdfReader(pdf_file_obj)
            for page_num in range(len(pdf_reader.pages)):
                page_obj = pdf_reader.pages[page_num]
                text += page_obj.extract_text()
    except FileNotFoundError:
        return "Error: PDF file not found."
    except Exception as e:
        return f"Error reading PDF: {e}"
    return text

def clean_text(text, newline, hyphen, doublespace, paragraph):
    """Removes extraneous newlines and other potential artifacts."""
    # Replace multiple newlines with a single space
    # This helps join lines that were broken for PDF formatting
    modifier = ""
    if newline:
        text = re.sub(r'\n+', ' ', text)
        modifier += "1"

    # Optional: Remove hyphenation at the end of lines
    if hyphen:
        text = re.sub(r'-\s+', '', text) # Uncomment if you have many hyphenated words
        modifier += "2"

    # Optional: Replace multiple spaces with a single space
    if doublespace:
        text = re.sub(r'\s{2,}', ' ', text)
        modifier += "3"

    # Attempt to re-insert paragraph breaks where they might have been.
    # This is a heuristic and might not be perfect.
    # It looks for a period followed by a space and an uppercase letter (start of a new sentence)
    # and assumes this might have been a paragraph break.
    # You might need to adjust this or remove it if it doesn't work well for your PDFs.
    if paragraph:
        text = re.sub(r'(?<=[.!?])\s+(?=[A-Z])', '\n\n', text)
        modifier += "4"

    return text.strip(), modifier

def save_text_to_word(text, word_path):
    """Saves the given text to a Word document."""
    try:
        document = Document()
        document.add_paragraph(text)
        document.save(word_path)
        return f"Successfully saved to {word_path}"
    except Exception as e:
        return f"Error saving Word document: {e}"

if __name__ == "__main__":
    pdf_input_path = input("Enter the path to your PDF file: ")
    word_output_path = input("Enter the desired path for the output Word file (e.g., output.docx): ")

    print("\nProcessing...")

    # 1. Extract text from PDF
    raw_text = extract_text_from_pdf(pdf_input_path)

    if "Error:" in raw_text:
        print(raw_text)
    else:
        # 2. Clean the extracted text
        print("Text extracted, now cleaning...")
        cleaned_text = clean_text(raw_text)

        # 3. Save cleaned text to Word
        save_message = save_text_to_word(cleaned_text, word_output_path)
        print(save_message)

        if "Successfully" in save_message:
            print("\n✨ PDF content has been transferred to the Word document! ✨")
        else:
            print("\n⚠️ There was an issue during the process.")