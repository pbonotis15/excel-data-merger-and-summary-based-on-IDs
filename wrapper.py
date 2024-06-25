# wrapper.py

import __innit__
import tkinter as tk
from tkinter import filedialog, simpledialog

def get_file_path():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx;*.xls")])

    return file_path

def get_excel_sheet(file_path):
    # Read all sheet names from the Excel file
    sheet_names = __innit__.get_sheet_names(file_path)
    
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    sheet_name = simpledialog.askstring("Select Excel Sheet", f"Available sheets in {file_path}:\n{', '.join(sheet_names)}\nEnter the sheet name:")

    return sheet_name

def get_output_folder():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    folder_path = filedialog.askdirectory(title="Select Output Folder")

    return folder_path


if __name__ == "__main__":
    file_path1 = get_file_path()
    file_path2 = get_file_path()
    
    if file_path1:
        file_path_sheet = file_path1
        sheet_name = get_excel_sheet(file_path_sheet)
        output_folder = get_output_folder()

        if output_folder:
            __innit__.main(file_path1, file_path2, sheet_name, output_folder)  # Pass the selected file path to your __init__.py script
