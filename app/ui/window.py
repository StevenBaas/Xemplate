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
    root = ctk.CTk()
    root.title("Xemplate")
    root.protocol("WM_DELETE_WINDOW", root.quit)
    screen_height = root.winfo_screenheight()
    screen_width = root.winfo_screenwidth()
    root.geometry(f"{screen_height}x{int(screen_width/2)}+0+0")

    excel_group = ctk.CTkFrame(master=root)
    excel_group.pack(pady=10)
    word_group = ctk.CTkFrame(master=root)
    word_group.pack(pady=10)

    excel_upload_button = ctk.CTkButton(master=excel_group, text="Upload Excel Data", command=open_excel_file)
    excel_upload_button.pack(pady=10)
    excel_sheet_name_label = ctk.CTkLabel(master=excel_group, text="Excel Sheet Name:")
    excel_sheet_name_label.pack(pady=10)
    excel_sheet_name_text = ctk.CTkEntry(master=excel_group, width=30)
    excel_sheet_name_text.pack(pady=10)

    word_upload_button = ctk.CTkButton(master=word_group, text="Upload Word Template", command=open_word_file)
    word_upload_button.pack(pady=10)
    output_file_name_label = ctk.CTkLabel(master=word_group, text="Output File Name:")
    output_file_name_label.pack(pady=10)
    output_file_name_text = ctk.CTkEntry(master=word_group, width=30)
    output_file_name_text.pack(pady=10)

    ctk.CTkButton(master=root, text="Propogate Template", command=call_propogate_template).pack()

    root.mainloop()
