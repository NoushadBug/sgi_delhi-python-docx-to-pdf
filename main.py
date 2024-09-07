import os
import shutil
from datetime import datetime
import time
from docx import Document
from deep_translator import GoogleTranslator
import win32com.client

def get_folder_path(args):
    try:
        if args:
            return args[0]
        else:
            return input("Enter the directory path containing the text files: ")
    except Exception as e:
        print(f"Error retrieving folder path: {e}")
        exit(1)

def merge_text_files(folder_path):
    try:
        combined_text = ""
        for file_name in os.listdir(folder_path):
            if file_name.endswith('.txt'):
                file_path = os.path.join(folder_path, file_name)
                with open(file_path, 'r', encoding='utf-8') as file:
                    lines = file.readlines()
                    if lines:
                        first_line = lines[0].strip()  # First line
                        remaining_lines = ''.join(line.strip() for line in lines[1:])  # Remaining lines
                        combined_text += f"{first_line}\n{remaining_lines}\n\n"
        return combined_text.strip()
    except Exception as e:
        print(f"Error merging text files: {e}")
        exit(1)

def split_text(text, max_length=5000):
    chunks = []
    while len(text) > max_length:
        split_index = text.rfind('.', 0, max_length)
        if split_index == -1:
            split_index = max_length
        chunks.append(text[:split_index+1])
        text = text[split_index+1:].strip()
    chunks.append(text)
    return chunks

def translate_text(text_chunks):
    try:
        translator = GoogleTranslator(source='auto', target='en')
        translated_chunks = []
        for chunk in text_chunks:
            translated = translator.translate(chunk)
            translated_chunks.append(translated)
        return ' '.join(translated_chunks)
    except Exception as e:
        print(f"Error translating text: {e}")
        exit(1)

def create_docx_with_translated_text(template_path, output_path, translated_text):
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_docx_path = os.path.join(output_path, f"{timestamp}.docx")
        shutil.copy(template_path, output_docx_path)

        doc = Document(output_docx_path)
        for paragraph in doc.paragraphs:
            if "{{combinedText}}" in paragraph.text:
                paragraph.text = paragraph.text.replace("{{combinedText}}", translated_text)
                paragraph.style = paragraph.style

        # Apply alignment and justification, preserving specific headlines' alignments
        for para in doc.paragraphs:
            text_content = para.text.strip()
            if text_content and text_content.split('\n')[0] == translated_text.split('\n')[0]:
                para.alignment = 0  # Left alignment
            elif text_content in ["Disclaimer", "OCR/HTR"]:
                continue  # Preserve original alignment
            else:
                para.alignment = 3  # Justified alignment
        
        doc.save(output_docx_path)
        return output_docx_path
    except Exception as e:
        print(f"Error creating DOCX with translated text: {e}")
        exit(1)

def convert_docx_to_pdf(docx_path):
    try:
        docx_path = os.path.abspath(docx_path)
        print(f"Full path for DOCX: {docx_path}")

        if not os.path.isfile(docx_path):
            raise FileNotFoundError(f"The file {docx_path} does not exist.")

        time.sleep(1)

        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(docx_path)

        pdf_path = docx_path.replace(".docx", ".pdf")
        print(f"Converting to PDF: {pdf_path}")

        doc.SaveAs(pdf_path, FileFormat=17)
        doc.Close()
        word.Quit()

        return pdf_path
    except Exception as e:
        print(f"Error converting DOCX to PDF: {e}")
        exit(1)

def clean_up(file_path):
    try:
        os.remove(file_path)
    except Exception as e:
        print(f"Error cleaning up file: {e}")

def main(args=None):
    try:
        folder_path = get_folder_path(args)
        template_path = './template/combiner-template.docx'
        output_path = './output/'

        os.makedirs(output_path, exist_ok=True)

        combined_text = merge_text_files(folder_path)
        text_chunks = split_text(combined_text)
        translated_text = translate_text(text_chunks)
        output_docx_path = create_docx_with_translated_text(template_path, output_path, translated_text)
        pdf_path = convert_docx_to_pdf(output_docx_path)
        clean_up(output_docx_path)

        print(f"PDF created successfully: {pdf_path}")

    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    import sys
    main(sys.argv[1:])
