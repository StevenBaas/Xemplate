import customtkinter as ctk
from tkinter import filedialog, messagebox
from app.state.app_state import app_state
from app.logic.data_processing import *


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
    def add_placeholder_and_replacement():
        placeholder_and_replacement_group = ctk.CTkFrame(master=right_group)
        placeholder_and_replacement_group.pack(pady=5)

        placeholder_label = ctk.CTkLabel(master=placeholder_and_replacement_group, text="Placeholder:")
        placeholder_label.pack(pady=5)
        placeholder_entry = ctk.CTkEntry(master=placeholder_and_replacement_group, width=100)
        placeholder_entry.pack(pady=5)

        replacement_label = ctk.CTkLabel(master=placeholder_and_replacement_group, text="Replacement Column:")
        replacement_label.pack(pady=5)
        replacement_entry = ctk.CTkEntry(master=placeholder_and_replacement_group, width=100)
        replacement_entry.pack(pady=5)


    root = ctk.CTk()
    root.title("Xemplate")
    root.protocol("WM_DELETE_WINDOW", root.quit)
    screen_height = root.winfo_screenheight()
    screen_width = root.winfo_screenwidth()
    root.geometry(f"{screen_height}x{int(screen_width/2)}+0+0")

    left_group = ctk.CTkFrame(master=root)
    left_group.pack(pady=5, side="left")
    excel_group = ctk.CTkFrame(master=left_group)
    excel_group.pack(pady=5)
    word_group = ctk.CTkFrame(master=left_group)
    word_group.pack(pady=5)
    right_group = ctk.CTkFrame(master=root)
    right_group.pack(pady=5, side="right")

    excel_upload_button = ctk.CTkButton(master=excel_group, text="Upload Excel Data", command=open_excel_file)
    excel_upload_button.pack(pady=5)
    excel_sheet_name_label = ctk.CTkLabel(master=excel_group, text="Excel Sheet Name:")
    excel_sheet_name_label.pack(pady=5)
    excel_sheet_name_text = ctk.CTkEntry(master=excel_group, width=100)
    excel_sheet_name_text.pack(pady=5)

    word_upload_button = ctk.CTkButton(master=word_group, text="Upload Word Template", command=open_word_file)
    word_upload_button.pack(pady=5)

    output_file_name_label = ctk.CTkLabel(master=word_group, text="Output File Name:")
    output_file_name_label.pack(pady=5)
    output_file_name_text = ctk.CTkEntry(master=word_group, width=100)
    output_file_name_text.pack(pady=5)

    placeholder_label = ctk.CTkLabel(master=right_group, text="Placeholder:")
    placeholder_label.pack(pady=5)
    placeholder_entry = ctk.CTkEntry(master=right_group, width=100)
    placeholder_entry.pack(pady=5)

    replacement_label = ctk.CTkLabel(master=right_group, text="Replacement Column:")
    replacement_label.pack(pady=5)
    replacement_entry = ctk.CTkEntry(master=right_group, width=100)
    replacement_entry.pack(pady=5)

    add_placeholder_and_replacement_button = ctk.CTkButton(master=root, text="Add Placeholder and Replacement", command=add_placeholder_and_replacement)
    add_placeholder_and_replacement_button.pack(pady=5)

    ctk.CTkButton(master=root, text="Propogate Template", command=call_propogate_template).pack()

    root.mainloop()
