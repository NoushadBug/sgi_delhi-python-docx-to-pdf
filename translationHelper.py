
# Define a list of sentence-ending characters for Urdu, Hindi, Gujarati, and other mentioned languages
import re
from deep_translator import GoogleTranslator
from langdetect import detect, DetectorFactory
from concurrent.futures import ThreadPoolExecutor, as_completed

DetectorFactory.seed = 0  # Ensure reproducible results

SENTENCE_ENDING_CHARACTERS = ['.', '!', '?', '۔', '؟', '।', '॥', '؛']
# Combine the sentence-ending characters into a regex pattern for splitting the text
SENTENCE_SPLIT_REGEX = r'(?<=[' + ''.join(re.escape(c) for c in SENTENCE_ENDING_CHARACTERS) + r'])\s+'

def is_english_text(sentence):
    # Filter the sentence to check for English letters, digits, and valid superscript/subscript characters
    english_chars = sum(1 for char in sentence if (char.isascii() and char.isalnum()) or char in "¹²³⁰⁴⁵⁶⁷⁸⁹₀₁₂₃₄₅₆₇₈₉")
    total_chars = sum(1 for char in sentence if char.isalnum() or char in "¹²³⁰⁴⁵⁶⁷⁸⁹₀₁₂₃₄₅₆₇₈₉")

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
