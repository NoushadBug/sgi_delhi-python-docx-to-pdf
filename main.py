import os, sys, win32com.client, json
from datetime import datetime
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from langdetect import detect, DetectorFactory
from concurrent.futures import ThreadPoolExecutor, as_completed

DetectorFactory.seed = 0  # Ensure reproducible results

def set_default_font(doc, font_name="Noto Sans"):
    """Set the default font for the entire document."""
    style = doc.styles['Normal']
    font = style.font
    font.name = font_name
    # Handle the case where the font needs to be set specifically for East Asian languages
    font.element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

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
                        text_content = ''.join(content[1:])  # Remaining text

                        # Left-align the first line (page heading) and make it bold
                        add_paragraph(doc, page_heading, WD_ALIGN_PARAGRAPH.LEFT, item['fontSize'] + 2, bold=True)
                        # Justify the rest of the text
                        add_paragraph(doc, text_content, WD_ALIGN_PARAGRAPH.JUSTIFY, item['fontSize'])
                        # Add a newline after each file's content
                        doc.add_paragraph()

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
