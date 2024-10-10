import os, json, sys, win32com.client

if os.name == 'nt':
    import msvcrt
else:
    import tty
    import termios
    
SUCCESS_INDICATOR = "    ✔️ "
ERROR_INDICATOR = "    ❌ "

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

def copy_directory(src, dst):
    # Extract the folder name from the source path
    folder_name = os.path.basename(src)
    dst_folder = os.path.join(dst, folder_name)

    # Ensure the destination folder exists
    os.makedirs(dst_folder, exist_ok=True)

    # Iterate over all items in the source directory
    for item in os.listdir(src):
        src_item = os.path.join(src, item)
        dst_item = os.path.join(dst_folder, item)

        # Copy file
        if os.path.isfile(src_item):
            print(f"Copying file: {src_item} to {dst_item}")
            with open(src_item, 'rb') as fsrc:
                with open(dst_item, 'wb') as fdst:
                    fdst.write(fsrc.read())

        # Copy directory (create the subdirectory and copy contents)
        elif os.path.isdir(src_item):
            print(f"Copying directory: {src_item} to {dst_item}")
            os.makedirs(dst_item, exist_ok=True)
            copy_directory(src_item, dst_item)  # Recursive call for subdirectory

def change_directory(config, config_path, key):
    try:
        new_directory = input(f"Enter the new {key.replace('_', ' ')} path: ")
        if key in config:  # Ensure key exists in config
            config[key] = new_directory
            print(f"Updating {key}: {config[key]}")
            save_json_config(config_path, config)
            print(f"{SUCCESS_INDICATOR} {key.replace('_', ' ').title()} updated to: {new_directory}")
        else:
            print(f"{ERROR_INDICATOR} Key '{key}' not found in the config.")
    except Exception as e:
        print(f"{ERROR_INDICATOR} Error updating {key.replace('_', ' ').title()}: {str(e)}")


def display_menu():
    clear_terminal()
    print("\nMain Menu:")
    print("    #1. Run Script")
    print("    #2. Set Input Folder Directory")
    print("    #3. Set Output Folder Directory")
    print("    #4. Set Copy Destination Folder")  # New option added
    print("    #5. Exit")
    
    print("\nEnter your choice (1/2/3/4/5): ", end='', flush=True)
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
        print(f"{ERROR_INDICATOR}Error converting DOCX to PDF: {e}")
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
