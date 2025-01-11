
# Georgian Stores Python Project

## Overview

The **Georgian Stores Python Project** is a Python application designed to process Excel files row by row, isolate and count types of resources based on a user-provided set of keywords. The application generates a report itemized by keywords, showing how many rows match each keyword. It ensures that each keyword is counted only once per row, even if it appears multiple times.

The end-user interacts with a graphical interface to:
1. Upload an Excel file for processing.
2. Provide a list of keywords to search for in the file.
3. Run the report to analyze the data.
4. Save the report to an output Excel file that appends the results every time a new file is uploaded.

## Features
- **GUI Interface**: Users can upload Excel files and specify keywords via a user-friendly interface.
- **Keyword Matching**: Counts keyword occurrences per row while ensuring each keyword is counted only once per line.
- **Automated Report Generation**: Processes the file and appends statistics as a new row in the output Excel file.

## Folder Structure

```
Georgian-Stores-Python-Project/
│
├── main.py                   # Entry point for the application
├── README.md                 # Project documentation
├── .gitignore                # Ignore files for version control
│
├── interface/                # User interface logic
│   └── user_interface.py     # Code for the GUI
│
├── processing/               # Core processing logic
│   └── excel_processor.py    # Processes Excel files and counts keyword occurrences
│
├── file_management/          # File input/output management
│   ├── file_handler.py       # Handles file uploads and saves
│   └── data/                 # Folder for input/output files
│       ├── input/            # User-uploaded files
│       └── output/           # Processed files (new or updated Excel files)
```

## Requirements
To run the application, ensure the following dependencies are installed:
- **Python 3.8 or higher**
- **Tkinter** (for the graphical user interface)
- **openpyxl** (to handle Excel file creation)

## How It Works
1. **User Uploads an Excel File**:
   - The user selects an Excel file through the graphical interface.
2. **Provide Keywords**:
   - The user inputs a list of keywords (e.g., "Digital", "Online Code") for the application to search for.
3. **Run the Report**:
   - The application processes the Excel file row by row, counts keyword matches, and appends the results to an output Excel file.
4. **Save the Report**:
   - The generated report is saved in the `output/` folder under `file_management/data/`.

## How to Run the Application
1. Clone this repository or download the files.
2. Ensure Python is installed on your system.
3. Navigate to the project directory and run:
   ```bash
   pip install openpyxl
   python main.py
   ```
4. Follow the on-screen instructions to upload an Excel file, provide keywords, and process the report.

## Example Use Case
### Input Excel File
| Row | Data                   |
|-----|------------------------|


### User Keywords
- Digital
- Online Code

### Output (Appended Row in Output File)
| Keyword        | Count |
|----------------|-------|


At the end you just have to install pyinputinstaller and convert this app to an executable file 


