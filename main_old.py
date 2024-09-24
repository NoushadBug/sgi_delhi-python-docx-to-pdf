import os, win32com.client, re, time, shutil
from datetime import datetime
from docx import Document
from deep_translator import GoogleTranslator
from concurrent.futures import ThreadPoolExecutor, as_completed
from langdetect import detect, DetectorFactory
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
DetectorFactory.seed = 0  # Ensure reproducible results

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
    combined_text = ""
    try:
        for file_name in os.listdir(folder_path):
            if file_name.endswith('.txt'):
                file_path = os.path.join(folder_path, file_name)
                with open(file_path, 'r', encoding='utf-8') as file:
                    lines = file.readlines()
                    if lines:
                        # Extract the first line and the rest of the lines
                        first_line = lines[0].strip()
                        remaining_lines = ''.join(line.strip() for line in lines[1:])
                        print(first_line)
                        # Combine the first line with the formatted remaining lines
                        combined_text += f"##LEFT##{first_line}##LEFT##\n{remaining_lines}\n\n"
        return combined_text.strip()
    except Exception as e:
        print(f"Error merging text files: {e}")
        exit(1)

def remove_sentence_ending_characters(text):
    """
    Remove all sentence-ending characters from the given text.
    """
    # Create a translation table to remove sentence-ending characters
    translation_table = str.maketrans('', '', ''.join(SENTENCE_ENDING_CHARACTERS))
    return text.translate(translation_table)

def is_sentence_numeric(sentence):
    """
    Check if a sentence is numeric after removing all sentence-ending characters.
    """
    cleaned_sentence = remove_sentence_ending_characters(sentence)
    return cleaned_sentence.isnumeric()

# Define a list of sentence-ending characters for the languages mentioned
SENTENCE_ENDING_CHARACTERS = ['.', '!', '?', '۔', '؟', '।', '॥', '؛']

def split_text(text, max_character=5000):
    # Create a regex pattern to split text based on sentence-ending characters
    split_pattern = f"(?<=[{''.join(SENTENCE_ENDING_CHARACTERS)}])(?=\s)"
    
    # Split text based on the pattern
    sentences = re.split(split_pattern, text)
    
    # Remove any empty strings from the list and strip whitespace
    # sentences = [sentence.strip() for sentence in sentences if sentence.strip()]
    
    grouped_sentences = []
    current_group = []
    current_lang = None
    current_length = 0
    detected_languages = []  # List to keep track of detected languages for each group

    for sentence in sentences:
        # Skip detection for numeric or trivial content
        if is_sentence_numeric(sentence):
            detected_lang = current_lang  # Assume it matches the current group if possible
        else:
            try:
                detected_lang = detect(sentence)
            except Exception as e:
                print(f"Error detecting language for: {sentence}. Error: {e}")
                detected_lang = None

        # Check if the sentence can be added to the current group without exceeding max_character
        if detected_lang == current_lang and (current_length + len(sentence) + 1) <= max_character:
            # If the language matches and adding the sentence does not exceed max_character
            current_group.append(sentence)
            current_length += len(sentence) + 1  # +1 for the space or punctuation that joins sentences
        else:
            # If there is a change in language or adding the sentence exceeds max_character
            if current_group:
                # Add the current group to the list if it's not empty
                grouped_sentences.append(' '.join(current_group))
                detected_languages.append(current_lang)  # Append the detected language for the current group
            # Start a new group with the new language
            current_group = [sentence]
            current_lang = detected_lang
            current_length = len(sentence) + 1

    # Append the last group if it exists
    if current_group:
        grouped_sentences.append(' '.join(current_group))
        detected_languages.append(current_lang)  # Append the detected language for the last group

    return grouped_sentences, detected_languages

def batch_translate(groups, langs):
    """
    Batch translate the given groups of text based on their languages.
    """
    translator = GoogleTranslator(source='auto', target='en')
    translated_chunks = []

    for group, lang in zip(groups, langs):
        if lang != 'en':  # Only translate non-English text
            # print(lang)
            # print(group)
            translated = translator.translate(group)
            translated_chunks.append(translated)
        else:
            translated_chunks.append(group)  # Directly append English text without translation

    return ' '.join(translated_chunks)

def parallel_process(groups, langs):
    """
    Use parallel threads to process translation faster.
    """
    with ThreadPoolExecutor(max_workers=4) as executor:  # Adjust max_workers based on your CPU cores
        future_to_translation = {executor.submit(batch_translate, [group], [lang]): group for group, lang in zip(groups, langs)}

        results = []
        for future in as_completed(future_to_translation):
            results.append(future.result())

    return ' '.join(results)

def create_docx_with_translated_text(template_path, output_path, translated_text):
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_docx_path = os.path.join(output_path, f"{timestamp}.docx")
        shutil.copy(template_path, output_docx_path)

        doc = Document(output_docx_path)
        
        # Replace placeholder with translated text
        for paragraph in doc.paragraphs:
            if "{{combinedText}}" in paragraph.text:
                # Use a placeholder to keep the formatting intact
                for run in paragraph.runs:
                    if "{{combinedText}}" in run.text:
                        run.text = run.text.replace("{{combinedText}}", translated_text)

        # Align text to the left when enclosed in ##LEFT##...##LEFT##
        for paragraph in doc.paragraphs:
            if '##LEFT##' in paragraph.text:
                # Extract the text inside ##LEFT## tags and remove the tags
                new_text = paragraph.text.replace('##LEFT##', '').strip()
                paragraph.text = new_text  # Set the new text without the tags
                
                # # Set paragraph alignment to left
                # paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

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
        # Step 1: Split and detect language
        grouped_sentences, detected_languages = split_text(combined_text)
        print(f"Processing total: {len(grouped_sentences)} sentences")
        # Step 2: Parallel translate
        translated_text = parallel_process(grouped_sentences, detected_languages)
        print(f"Translating process done..")
        output_docx_path = create_docx_with_translated_text(template_path, output_path, translated_text)
        pdf_path = convert_docx_to_pdf(output_docx_path)
        clean_up(output_docx_path)

        print(f"PDF created successfully: {pdf_path}")

    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    import sys
    main(sys.argv[1:])
