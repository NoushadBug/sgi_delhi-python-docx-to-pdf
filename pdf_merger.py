import os
import sys
from PyPDF2 import PdfMerger

def merge_pdfs(directory):
    try:
        # Get all PDF files in the directory, sorted by name
        pdf_files = [f for f in os.listdir(directory) if f.lower().endswith('.pdf') and f != 'merged_output.pdf']
        pdf_files.sort()

        if not pdf_files:
            print("No PDF files found in the directory.")
            return

        # Initialize the PdfMerger object
        merger = PdfMerger()

        # Merge all the sorted PDF files
        for pdf in pdf_files:
            pdf_path = os.path.join(directory, pdf)
            merger.append(pdf_path)
            print(f"Merging: {pdf}")

        # Save the merged PDF to the same directory
        output_file = os.path.join(directory, 'merged_output.pdf')
        merger.write(output_file)
        merger.close()

        print(f"All PDFs have been merged into: {output_file}")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    # Check if directory argument is provided, else ask the user
    if len(sys.argv) > 1:
        input_directory = sys.argv[1]
    else:
        input_directory = input("Please enter the directory containing PDF files: ")

    # Check if the provided directory is valid
    if not os.path.isdir(input_directory):
        print("The provided directory is not valid.")
    else:
        merge_pdfs(input_directory)
