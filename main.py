import os, sys, win32com.client, json, re
from datetime import datetime
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from deep_translator import GoogleTranslator
from langdetect import detect, DetectorFactory
from concurrent.futures import ThreadPoolExecutor, as_completed
from langdetect.lang_detect_exception import LangDetectException

DetectorFactory.seed = 0  # Ensure reproducible results

def set_default_font(doc, font_name="Noto Sans"):
    """Set the default font for the entire document."""
    style = doc.styles['Normal']
    font = style.font
    font.name = font_name
    # Handle the case where the font needs to be set specifically for East Asian languages
    font.element.rPr.rFonts.set(qn('w:eastAsia'), font_name)


# Define a list of sentence-ending characters for Urdu, Hindi, Gujarati, and other mentioned languages
SENTENCE_ENDING_CHARACTERS = ['.', '!', '?', '۔', '؟', '।', '॥', '؛']

# Combine the sentence-ending characters into a regex pattern for splitting the text
SENTENCE_SPLIT_REGEX = r'(?<=[' + ''.join(re.escape(c) for c in SENTENCE_ENDING_CHARACTERS) + r'])\s+'

def is_english_text(sentence):
    # Filter the sentence to check for English letters and digits
    english_chars = sum(1 for char in sentence if char.isascii() and char.isalnum())
    total_chars = sum(1 for char in sentence if char.isalnum())
    
    # If most of the alphanumeric characters are ASCII, assume it's English
    if total_chars == 0:
        return True  # If no alphanumeric characters, treat as English (neutral).
    
    return english_chars / total_chars > 0.5  # More than 50% English characters.


def translate_chunk(chunk):
    translator = GoogleTranslator(source='auto', target='en')
    return translator.translate(chunk.strip())

def translate_to_english(text, chunk_size=5000, max_workers=5):
    # Break text into chunks that are under the limit
    chunks = []
    current_chunk = []

    current_length = 0
    for sentence in text.split('.'):  # Split text by sentences to ensure coherent chunks
        sentence_length = len(sentence) + 1  # Account for the period
        if current_length + sentence_length <= chunk_size:
            current_chunk.append(sentence)
            current_length += sentence_length
        else:
            chunks.append('.'.join(current_chunk) + '.')
            current_chunk = [sentence]
            current_length = sentence_length

    # Add the last chunk
    if current_chunk:
        chunks.append('.'.join(current_chunk) + '.')

    # Use threading to translate chunks in parallel
    translated_chunks = []
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        # Submit all translation tasks
        futures = {executor.submit(translate_chunk, chunk): chunk for chunk in chunks}

        # Collect results as they complete
        for future in as_completed(futures):
            translated_chunks.append(future.result())

    # Join all the translated chunks into the final translated text
    translated_text = ' '.join(translated_chunks)
    return translated_text

def check_and_translate_non_english_content(text_content, max_workers=5):
    # Split the content using sentence-ending characters while preserving the delimiters
    sentences = re.split(SENTENCE_SPLIT_REGEX, text_content)

    contains_non_english = False
    translated_text = [None] * len(sentences)  # Initialize a list to store the translated sentences

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        # Prepare futures for language detection with their original index
        language_futures = {executor.submit(is_english_text, sentence): (i, sentence) for i, sentence in enumerate(sentences) if len(sentence.strip()) > 3}

        for future in as_completed(language_futures):
            index, sentence = language_futures[future]
            is_english = future.result()

            if not is_english:  # Non-English content detected
                translated_sentence = translate_to_english(sentence)
                contains_non_english = True
            else:
                translated_sentence = sentence

            # Store the translated sentence in its original index
            translated_text[index] = translated_sentence + " "

    # Properly join sentences, ensuring spaces between them and preserving punctuation
    final_text = ''
    for sentence in translated_text:
        if sentence:  # Make sure to skip None values
            final_text += sentence
            # Add a space if the last character is not a punctuation mark
            if not sentence[-1] in SENTENCE_ENDING_CHARACTERS:
                final_text += ' '

    return contains_non_english, final_text.strip()

def create_docx_from_structure(output_path, structure, text_files):
    try:
        # Generate timestamp and output paths
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_docx_path = os.path.join(output_path, f"{timestamp}.docx")

        # Create a new DOCX document
        doc = Document()

        # Set the default font to Noto Sans
        set_default_font(doc)

        # Helper function to add paragraphs
        def add_paragraph(doc, text, alignment, font_size, bold=False):
            paragraph = doc.add_paragraph()
            run = paragraph.add_run(text)
            run.bold = bold
            run.font.size = Pt(font_size)
            paragraph.alignment = alignment

        # Accumulate all translated content to add at the end
        translation_section = []

        # Process each item in the structure list
        for item in structure:
            if item['type'] == 'heading':
                add_paragraph(doc, item['text'], WD_ALIGN_PARAGRAPH.CENTER, item['fontSize'], bold=True)
            elif item['type'] == 'normal':
                add_paragraph(doc, item['text'], WD_ALIGN_PARAGRAPH.JUSTIFY, item['fontSize'])
            elif item['type'] == 'pagebreak':
                doc.add_page_break()
            elif item['type'] == 'combinedText':
                # Insert the content from each text file
                for text_file in text_files:
                    with open(text_file, 'r', encoding='utf-8') as f:
                        content = f.readlines()
                        page_heading = content[0].strip()  # First line (page heading)
                        # Replace spaces with hyphens
                        page_heading = page_heading.replace(" ", "-")
                        text_content = ''.join(content[1:])  # Remaining text

                        # Remove leading newline characters if they exist
                        if text_content.startswith('\n'):
                            text_content = text_content.lstrip('\n')

                        # Check for non-English content and translate it immediately
                        contains_non_english, translated_text = check_and_translate_non_english_content(text_content)

                        # Left-align the first line (page heading) and make it bold
                        add_paragraph(doc, page_heading, WD_ALIGN_PARAGRAPH.LEFT, item['fontSize'] + 2, bold=True)
                        # Justify the rest of the text
                        add_paragraph(doc, text_content, WD_ALIGN_PARAGRAPH.JUSTIFY, item['fontSize'])

                        # If non-English content is detected, accumulate the translation
                        if contains_non_english:
                            translation_section.append({
                                'heading': page_heading,
                                'translated_text': translated_text,
                                'font_size': item['fontSize']
                            })

                        # Add a newline after each file's content
                        doc.add_paragraph()

        # Add the translation section at the end of the document
        if translation_section:
            add_paragraph(doc, "Translation", WD_ALIGN_PARAGRAPH.CENTER, item['fontSize'] + 4, bold=True)
            for section in translation_section:
                add_paragraph(doc, section['heading'], WD_ALIGN_PARAGRAPH.LEFT, section['font_size'] + 2, bold=True)
                add_paragraph(doc, section['translated_text'], WD_ALIGN_PARAGRAPH.JUSTIFY, section['font_size'])
                add_paragraph(doc, "", WD_ALIGN_PARAGRAPH.JUSTIFY, section['font_size'])

        # Save the DOCX file
        doc.save(output_docx_path)
        return output_docx_path
    except Exception as e:
        print(f"Error creating DOCX: {e}")
        exit(1)

def convert_docx_to_pdf(docx_path):
    try:
        docx_path = os.path.abspath(docx_path)
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(docx_path)

        pdf_path = docx_path.replace(".docx", ".pdf")
        doc.SaveAs(pdf_path, FileFormat=17)  # 17 is the format code for PDF
        doc.Close()
        word.Quit()

        # Delete the DOCX file after conversion
        os.remove(docx_path)
        return pdf_path
    except Exception as e:
        print(f"Error converting DOCX to PDF: {e}")
        exit(1)

def get_folder_path(args=None):
    # Check if a folder path was passed as an argument
    if args and len(args) > 1:
        folder_path = args[1]  # Get the second argument (first is script name)
    else:
        # Ask for input if no argument provided
        folder_path = input("Please enter the folder path for text files: ").strip()

    # Validate the folder path
    if not os.path.isdir(folder_path):
        print(f"Invalid folder path: {folder_path}")
        exit(1)
    
    return folder_path

def load_json_config(config_file):
    """Load JSON configuration from a file."""
    with open(config_file, 'r') as file:
        return json.load(file)

def main(args=None):
    folder_path = get_folder_path(args)
    output_path = './output/'
    os.makedirs(output_path, exist_ok=True)

    # Load the document structure from config.json
    config = load_json_config('config.json')
    doc_structure = config['doc_structure']

    # Collect all text files from the folder
    text_files = [os.path.join(folder_path, file) for file in os.listdir(folder_path) if file.endswith('.txt')]

    # Create DOCX from text files
    docx_file = create_docx_from_structure(output_path, doc_structure, text_files)
    print(f"DOCX created at: {docx_file}")

    # Convert the DOCX to PDF
    pdf_file = convert_docx_to_pdf(docx_file)
    print(f"PDF created at: {pdf_file}")

if __name__ == "__main__":
    main(sys.argv)
