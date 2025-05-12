import tkinter as tk
from tkinter import filedialog, messagebox
from app.state.app_state import app_state
from app.logic.functions import *


def open_excel_file():
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file_path:
        print(f"Excel file selected: {file_path}")
        app_state.excel_file_path = file_path


def open_word_file():
    file_path = filedialog.askopenfilename(
        title="Select Word Template Document",
        filetypes=[("Word files", "*.docx *.doc")]
    )
    if file_path:
        print(f"Word document selected: {file_path}")
        app_state.word_template_path = file_path


def call_propogate_template():
    propogate_template(app_state.excel_file_path, app_state.excel_sheet_name, app_state.word_template_path)


def launch_ui():
    root = tk.Tk()
    root.title("Xemplate")
    root.protocol("WM_DELETE_WINDOW", root.quit)
    root.state('zoomed')

    excel_group = tk.Frame(root, bd=2, relief=tk.GROOVE)
    excel_group.pack()
    word_group = tk.Frame(root, bd=2, relief=tk.GROOVE)
    word_group.pack()

    tk.Button(excel_group, text="Upload Excel Data", command=open_excel_file).pack()
    tk.Label(excel_group, text="Excel Sheet Name:").pack()
    tk.Text(excel_group, height=1, width=30).pack()

    tk.Button(word_group, text="Upload Word Template", command=open_word_file).pack()
    tk.Label(word_group, text="Output Document Name:").pack()
    tk.Text(word_group, height=1, width=30).pack()
    
    tk.Button(root, text="Propogate Template", command=call_propogate_template).pack()

    root.mainloop()