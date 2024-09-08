import re
from langdetect import detect, DetectorFactory

# Seed the detector for consistent results
DetectorFactory.seed = 0

SENTENCE_ENDING_CHARACTERS = ['.', '!', '?', '۔', '؟', '।']  # Add characters based on your requirements

def split_text(text):
    # Create a regex pattern to split text based on sentence-ending characters
    split_pattern = f"(?<=[{''.join(SENTENCE_ENDING_CHARACTERS)}])(?=\s)"
    
    # Split text based on the pattern
    sentences = re.split(split_pattern, text)
    
    # Remove any empty strings from the list and strip whitespace
    sentences = [sentence.strip() for sentence in sentences if sentence.strip()]
    
    return sentences

def detect_and_group_sentences(text):
    sentences = split_text(text)
    grouped_sentences = []
    current_group = []
    current_lang = None

    for sentence in sentences:
        try:
            detected_lang = detect(sentence)
        except Exception as e:
            print(f"Error detecting language for: {sentence}. Error: {e}")
            detected_lang = None

        if detected_lang == current_lang:
            # If the language matches the current group, add it to the current group
            current_group.append(sentence)
        else:
            # If there is a change in language or this is the first sentence
            if current_group:
                # Add the current group to the list if it's not empty
                grouped_sentences.append(' '.join(current_group))
            # Start a new group with the new language
            current_group = [sentence]
            current_lang = detected_lang

    # Append the last group if it exists
    if current_group:
        grouped_sentences.append(' '.join(current_group))

    return grouped_sentences

# Example usage:
text = """This is an English sentence. This is another English sentence! یہ ایک اردو جملہ ہے۔  यह एक हिंदी वाक्य है। And yet another English sentence."""

grouped_sentences = detect_and_group_sentences(text)

for i, group in enumerate(grouped_sentences, 1):
    print(f"Group {i}: {group}")
