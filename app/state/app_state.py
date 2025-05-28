class AppState:
    def __init__(self):
        self.document_name_entry = None # Entry for document name in the UI
        self.sheet_name_entry = None # Entry for sheet name in the UI

        self.naming_coumns = [] # List of rows to be used for naming the documents
        self.excel_file_path = ""
        self.word_template_path = ""

        self.replacement = ""
        self.placeholder = ""

        self.output_folder_path = "output"
        self.output_file_name = "" # Main name of the output file

        self.personalization_columns = {} # Name to personalize each document, will probaly be column numbers in the sheet
        self.personalization_name = ""

        # UI Elements
        self.root = None
        self.left_group = None
        self.right_group = None
        self.placeholder_and_replacement_group = None
        self.placeholder_entries = [] # List of placeholder entries in the UI
        self.replacement_entries = [] # List of replacement entries in the UI

        # UI Variables
        self.current_row = 0
        self.current_column = 0

        #self.max_columns = self.get_max_columns()

    def get_max_columns(self):
            self.root.update_idletasks()
            root_width = self.root.winfo_width()
            left_group_width = self.left_group.winfo_width()
            placeholder_and_replacement_width = self.placeholder_and_replacement_group.winfo_width()
            max_columns = int((root_width + left_group_width) / placeholder_and_replacement_width)
            return max_columns
    
    def get_document_name(self):
        if self.document_name_entry:
            print(self.document_name_entry.get())
            return self.document_name_entry.get()
        return ""
    
    def get_sheet_name(self):
        if self.sheet_name_entry:
            print(self.document_name_entry.get())
            return self.sheet_name_entry.get()
        return ""


app_state = AppState()
