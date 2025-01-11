from file_management.file_handler import FileHandler
from interface.user_interface import UI
from processing.excel_processor import ExcelProcessor

if __name__ == "__main__":

    # References of the methods
    
    excelProcessorObj = ExcelProcessor()
    # this save excel file function is must for the fileHandlerObj
    create_excel_file_function = excelProcessorObj.create_excel_file

    fileHandlerObj = FileHandler(create_excel_file_function)

    # I need this function in the UI class
    save_excel_fileFunction = fileHandlerObj.save_excel_file
    
    interface = UI(save_excel_fileFunction)
    

