import customtkinter as ctk
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
    root = ctk.Tk()
    root.title("Xemplate")
    root.protocol("WM_DELETE_WINDOW", root.quit)
    root.state('zoomed')

    excel_group = ctk.Frame(root, bd=2, relief=ctk.GROOVE)
    excel_group.pack()
    word_group = ctk.Frame(root, bd=2, relief=ctk.GROOVE)
    word_group.pack()

    excel_upload_button = ctk.Button(excel_group, text="Upload Excel Data", command=open_excel_file)
    excel_upload_button.pack()
    excel_sheet_name_label = ctk.Label(excel_group, text="Excel Sheet Name:")
    excel_sheet_name_label.pack()
    excel_sheet_name_text = ctk.Text(excel_group, height=1, width=30)
    excel_sheet_name_text.pack()

    word_upload_button = ctk.Button(word_group, text="Upload Word Template", command=open_word_file)
    word_upload_button.pack()
    output_file_name_label = ctk.Label(word_group, text="Output File Name:")
    output_file_name_label.pack()
    output_file_name_text = ctk.Text(word_group, height=1, width=30)
    output_file_name_text.pack()


    
    ctk.Button(root, text="Propogate Template", command=call_propogate_template).pack()

    root.mainloop()
