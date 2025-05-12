class AppState:
    def __init__(self):
        self.excel_file_path = None
        self.excel_sheet_name = None # Name of sheet to be used in the Excel file
        self.word_template_path = None
        self.output_folder_path = "output"
        self.output_file_name = None # Main name of the output file
        self.personalization_column = None # Name to personalize each document, will probaly be a column in the sheet

app_state = AppState()
