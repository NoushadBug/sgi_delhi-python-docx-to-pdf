import os, json, sys, win32com.client

if os.name == 'nt':
    import msvcrt
else:
    import tty
    import termios

# Function to capture a single key press without Enter
def get_keypress():
    if os.name == 'nt':  # For Windows
        return msvcrt.getch().decode('utf-8')
    else:  # For Unix-based systems (Linux, macOS)
        fd = sys.stdin.fileno()
        old_settings = termios.tcgetattr(fd)
        try:
            tty.setraw(sys.stdin.fileno())
            ch = sys.stdin.read(1)
        finally:
            termios.tcsetattr(fd, termios.TCSADRAIN, old_settings)
        return ch

def wait_for_keypress():
    print("\n\n##██████████████████ Press any key to return the Menu. ██████████████████##\n\n")
    get_keypress()

def load_config(config_path):
    with open(config_path, 'r') as f:
        return json.load(f)

def clear_terminal():
    os.system('cls' if os.name == 'nt' else 'clear')

def change_directory(config, config_path, isRoot=True):
    try:
        new_directory = input(f"Enter the new {'Input' if isRoot else 'Output'} directory path: ")
        if isRoot:
            config['root_folder_directory'] = new_directory
        else:
            config['output_folder_directory'] = new_directory
        save_json_config(config_path, config)
        print(f"     ✔️ {'Input' if isRoot else 'Output'} Directory updated to: {new_directory}")
    except Exception as e:
        print(f"     ❌ Error updating {'Input' if isRoot else 'Output'} Directory: {str(e)}")


def display_menu():
    clear_terminal()
    print("\nMain Menu:")
    print("    #1. Run Script")
    print("    #2. Set Input Folder Directory")
    print("    #3. Set Output Folder Directory")
    print("    #4. Exit")
    
    print("\nEnter your choice (1/2/3/4): ", end='', flush=True)
    return get_keypress()

# Load JSON configuration
def load_json_config(filepath):
    with open(filepath, 'r') as file:
        return json.load(file)

# Save JSON configuration
def save_json_config(filepath, data):
    with open(filepath, 'w') as file:
        json.dump(data, file, indent=4)


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
        print(f"    ❌ Error converting DOCX to PDF: {e}")
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
