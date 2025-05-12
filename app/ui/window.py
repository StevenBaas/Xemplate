import tkinter as tk
from tkinter import filedialog, messagebox
from app.logic.functions import *


def open_excel_file():
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file_path:
        print(f"Excel file selected: {file_path}")
        # TODO: Process the Excel file here


def open_word_file():
    file_path = filedialog.askopenfilename(
        title="Select Word Template Document",
        filetypes=[("Word files", "*.docx *.doc")]
    )
    if file_path:
        print(f"Word document selected: {file_path}")
        # TODO: Process the Word document here


def launch_ui():
    root = tk.Tk()
    root.title("Xemplate")
    root.geometry("300x200")

    button = tk.Button(root, text="Upload Excel Data", command=open_excel_file)
    button.pack()
    button = tk.Button(root, text="Upload Word Template", command=open_word_file)
    button.pack()
    button = tk.Button(root, text="Propogate Template", command=propogate_template)
    button.pack()

    root.mainloop()