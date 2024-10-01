import os
import sys
from PyPDF2 import PdfMerger

def merge_pdfs(directory, passed_pdf_path=None):
    try:
        # Get all PDF files in the directory
        pdf_files = [os.path.join(directory, f) for f in os.listdir(directory) if f.lower().endswith('.pdf')]

        # Ensure there's exactly one PDF in the directory
        if len(pdf_files) > 1:
            print("    ❌ The directory must contain at least one PDF file.")
            return

        # Add the passed PDF full path to the list, without moving it
        if passed_pdf_path and os.path.isfile(passed_pdf_path):
            pdf_files.append(passed_pdf_path)  # Add the full path of the passed PDF

        # Initialize the PdfMerger object
        merger = PdfMerger()

        # Merge all the sorted PDF files
        for pdf_path in pdf_files:
            merger.append(pdf_path)
            print(f"Merging: {pdf_path}")

        # Use the original PDF filename for the merged output
        merged_pdf_filename = os.path.basename(pdf_files[0])  # This is the original PDF name
        output_file = os.path.join(directory, merged_pdf_filename)
        
        # Save the merged PDF to the same directory
        merger.write(output_file)
        merger.close()

        print(f"    ✔️ PDF(s) successfully merged into: {output_file} \n")

    except IndexError:
        print("Error: Attempted to access a list index that does not exist.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    # Check if directory argument is provided, else ask the user
    if len(sys.argv) > 1:
        passed_file_directory = sys.argv[1]
    else:
        passed_file_directory = None

    input_directory = input("Please enter the directory containing not more than 1 PDF file: ")

    # Check if the provided directory is valid
    if not os.path.isdir(input_directory):
        print("The provided directory is not valid.")
    else:
        merge_pdfs(input_directory, passed_file_directory)
