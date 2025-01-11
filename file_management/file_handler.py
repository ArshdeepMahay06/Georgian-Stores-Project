# fileHanderNew

import os
from datetime import datetime
from tkinter import filedialog

class FileHandler:

    # Input and Output folders
    OUTPUT_FOLDER = "../file_management"



    def __init__(self, create_excel_file_function):
        self.create_excel_file = create_excel_file_function
        os.makedirs(self.OUTPUT_FOLDER, exist_ok=True)
        
    

    def save_excel_file(self,source_path,keywords):
        return self.create_excel_file(keywords, source_path)

