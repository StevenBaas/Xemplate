import customtkinter as ctk
from tkinter import filedialog, messagebox
from app.state.app_state import app_state


def add_frame(master, pady=5, padx=5):
    frame = ctk.CTkFrame(master=master)
    frame.pack(pady=pady, padx=padx)
    return frame

def add_label(master, text, pady=5, padx=5):
    label = ctk.CTkLabel(master=master, text=text)
    label.pack(pady=pady, padx=padx)
    return label

def add_entry(master, width=100, pady=5, padx=5):
    entry = ctk.CTkEntry(master=master, width=width)
    entry.pack(pady=pady, padx=padx)
    return entry

def add_placeholder_and_replacement(master):
        placeholder_and_replacement_group = ctk.CTkFrame(master=master)
        placeholder_and_replacement_group.pack(pady=5)

        placeholder_label = ctk.CTkLabel(master=placeholder_and_replacement_group, text="Placeholder:")
        placeholder_label.pack(pady=5)
        placeholder_entry = ctk.CTkEntry(master=placeholder_and_replacement_group, width=100)
        placeholder_entry.pack(pady=5)

        replacement_label = ctk.CTkLabel(master=placeholder_and_replacement_group, text="Replacement Column:")
        replacement_label.pack(pady=5)
        replacement_entry = ctk.CTkEntry(master=placeholder_and_replacement_group, width=100)
        replacement_entry.pack(pady=5)