class AppState:
    def __init__(self):
        self.excel_file_path = ""
        self.excel_sheet_name = "" # Name of sheet to be used in the Excel file
        self.word_template_path = ""

        self.replacement = ""
        self.placeholder = ""

        self.output_folder_path = "output"
        self.output_file_name = "" # Main name of the output file

        self.personalization_columns = {} # Name to personalize each document, will probaly be column numbers in the sheet
        self.personalization_name = ""

        # UI Elements
        placeholder_entries = [] # List of placeholder entries in the UI
        replacement_entries = [] # List of replacement entries in the UI

app_state = AppState()
