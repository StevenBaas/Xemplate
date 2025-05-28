import os
from datetime import datetime

import openpyxl
from docx import Document
from docx2pdf import convert

from app.state.app_state import app_state


# Function to replace placeholder in paragraphs
def replace_placeholder_in_paragraphs(doc, placeholder, replacement):
    for paragraph in doc.paragraphs:
        if placeholder in paragraph.text:
            paragraph.text = paragraph.text.replace(placeholder, replacement)


# Function to replace placeholder in tables
def replace_placeholder_in_tables(doc, placeholder, replacement):
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if placeholder in cell.text:
                    cell.text = cell.text.replace(placeholder, replacement)


# Function to replace placeholder in headers and footers
def replace_placeholder_in_headers_footers(doc, placeholder, replacement):
    for section in doc.sections:
        header = section.header
        for paragraph in header.paragraphs:
            if placeholder in paragraph.text:
                paragraph.text = paragraph.text.replace(placeholder, replacement)

        footer = section.footer
        for paragraph in footer.paragraphs:
            if placeholder in paragraph.text:
                paragraph.text = paragraph.text.replace(placeholder, replacement)


# Function to replace placeholder in footnotes and endnotes
def replace_placeholder_in_notes(doc, placeholder, replacement):
    for footnote in doc.footnotes:
        if placeholder in footnote.text:
            footnote.text = footnote.text.replace(placeholder, replacement)
    for endnote in doc.endnotes:
        if placeholder in endnote.text:
            endnote.text = endnote.text.replace(placeholder, replacement)


# Function to replace placeholder in entire document
def replace_placeholder_in_document(doc, placeholder, replacement):
    replace_placeholder_in_paragraphs(doc, placeholder, replacement)
    replace_placeholder_in_tables(doc, placeholder, replacement)
    replace_placeholder_in_headers_footers(doc, placeholder, replacement)
    replace_placeholder_in_notes(doc, placeholder, replacement)


def get_data_as_list(data):
    data_list = []
    for row in data.iter_rows(values_only=True):
        data_row = []
        for cell in row:
            if (cell is not None) and (cell != ""):
                if isinstance(cell, str):
                    data_row.append(cell.strip())
                else:
                    data_row.append(str(cell).strip())
            else:
                data_row.append(cell)
        data_list.append(data_row)

    return data_list


def is_empty(data):
    if (data is None) or (data == "None") or (data == ""):
        return True
    else:
        return False


def convert_to_pdf(output_folder_path, output_file_name):
    convert(os.path.join(output_folder_path, output_file_name))
    os.remove(os.path.join(output_folder_path, output_file_name))


def propogate_template(excel_file_path, template_file_path):
    wb = openpyxl.load_workbook("Data.xlsx", data_only=True, read_only=True) # Will be variable for Uploaded Excel Document
    
    excel_sheet_name = app_state.get_sheet_name()
    excel_sheet = get_data_as_list(wb[excel_sheet_name])
    doc_name = app_state.get_document_name()

    for row in excel_sheet:
        # Open the Word document
        doc = Document("Template.docx") # Will be variable for Uploaded Template Document

        date = datetime.today()

        placeholder = "[DATE]"
        replacement = date.strftime("%d %B %Y")

        replace_placeholder_in_document(doc, placeholder, replacement)

        for i in range(13):
            if i == 1:
                placeholder = "[LAST_NAME]"
                replacement = row[i]

                if is_empty(replacement):
                    replacement = "N/A"
                replace_placeholder_in_document(doc, placeholder, replacement)
            elif i == 2:
                placeholder = "[FIRST_NAME]"
                replacement = row[i]

                if is_empty(replacement):
                    replacement = "N/A"
                replace_placeholder_in_document(doc, placeholder, replacement)
            elif i == 3:
                placeholder = "[STUDENT_NUMBER]"
                replacement = row[i]

                if is_empty(replacement):
                    replacement = "N/A"
                replace_placeholder_in_document(doc, placeholder, replacement)
            elif i == 4:
                placeholder = "[ID]"
                replacement = row[i]

                if is_empty(replacement):
                    replacement = "N/A"
                replace_placeholder_in_document(doc, placeholder, replacement)
            elif i == 5:
                placeholder = "[INSTITUTION]"
                replacement = row[i]

                if is_empty(replacement):
                    replacement = "Other Debt"
                replace_placeholder_in_document(doc, placeholder, replacement)
            elif i == 6:
                placeholder = "[QUALIFICATION]"
                replacement = row[i]

                if is_empty(replacement):
                    replacement = "N/A"
                replace_placeholder_in_document(doc, placeholder, replacement)

        # Save the modified document
        doc.save(f"output-docs\\{row[1]} {row[2]}_{doc_name}_{date.year}.docx")
        convert_to_pdf(f"output-docs",
                       f"{row[1]} {row[2]}_{doc_name}_{date.year}.docx")