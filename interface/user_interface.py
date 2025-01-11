# new userInterface

import tkinter as tk
from tkinter import filedialog
import datetime

class UI:

    window = tk.Tk()
    window.geometry("500x350")

    try:
        window.iconbitmap("appIcon.ico")
    except:
        pass
    window.title("Excel Analyzer")
    
    timeNow = datetime.datetime.now()

    keywordsLabel = tk.Label(text= "Enter your Keywords (Seprated by comma) :")
    keywordsEntry = tk.Entry(width=40)

    excelFilePath = None
    selectedFileNameVar = tk.StringVar()
    selectedFileNameVar.set("File Not Selected Yet")
    selectedFileLabel = tk.Label(text= "Your selected file is : ",pady=10)
    selectedFileNameLabel = tk.Label(textvar = selectedFileNameVar)

    totalRecordsLabel = tk.Label(text="Total Records found are : ",pady=10)
    totalRecordsFound = tk.StringVar()
    totalRecordsFound.set("None")
    totalRecordsFoundLabel = tk.Label(textvar = totalRecordsFound)
    
    
    def __init__(self,save_excel_fileFunction):

        self.save_excel_fileFunction = save_excel_fileFunction
        
        self.welcomeLabel = self.generateWelcomeLabel(self.greet())
        
        self.selectFileButton = tk.Button(self.window, text="Select Excel File", pady= 5, command= self.selectFileCommand)
        self.generateFileButton = tk.Button(self.window, text="Click Here to Generate File", pady = 5, command= self.generateFileCommand)

        
        self.welcomeLabel.grid(row=0,column=0 , pady=4,sticky=tk.W, columnspan=2)
        self.selectFileButton.grid(row = 1, column = 0, pady=2,sticky=tk.E)
        self.selectFileButton.configure(width=20)
        self.selectedFileLabel.grid(row=2,column=0,pady=4,sticky=tk.W, padx=40)
        self.selectedFileNameLabel.grid(row=2, column=1, pady=2,sticky=tk.W)
        self.keywordsLabel.grid(row=3,column=0,pady=4,sticky=tk.E)
        self.keywordsEntry.grid(row=3, column=1,sticky=tk.E)
        self.generateFileButton.grid(row=4 ,columnspan=2, pady=10,sticky=tk.W, padx=40)
        self.generateFileButton.configure(width=40)
        self.totalRecordsLabel.grid(row=5,column=0,pady=4,sticky=tk.W, padx=40)
        self.totalRecordsFoundLabel.grid(row=5, column=1, pady=2,sticky=tk.W)
        self.window.mainloop()


    def greet(self):
        currHour = self.timeNow.hour
        if currHour >= 5 and currHour<=12:
            return "Good Morning"
        elif currHour > 12 and currHour<=16:
            return "Good Afternoon"
        elif (currHour > 16 and currHour<= 23) or (currHour <=0 and currHour>5):
            return "Good Evening"

    def generateWelcomeLabel(self,greetFuncVal=""):
        welcomeLabel = tk.Label(text="Hi there! {0}".format(greetFuncVal),pady=20,padx=80, font=("Helvetica", 18), justify="left")
        return welcomeLabel

    def selectFileCommand(self):
        file_path = filedialog.askopenfilename(title="Select a File", filetypes=[("Excel Files", "*.xlsx"), ("All files", "*.*")])
        if file_path:
            # here the file path is stored
            self.excelFilePath = file_path
            self.selectedFileName = file_path.split(r'/')[-1]
            self.selectedFileNameVar.set(self.selectedFileName)
            

    def generateFileCommand(self):
        try:
            if(not self.excelFilePath or not self.keywordsEntry.get()):
        
                if(not self.excelFilePath):
                    # give a message that select your excel file
                    self.totalRecordsFound.set("Please select excel file first.")

                elif(not self.keywordsEntry.get()):    
                    # give a message to enter keywords
                    self.totalRecordsFound.set("Please give some keywords first.")
                
                return
            keywords = self.keywordsEntry.get()
            # code to generate the excel file
            total = self.save_excel_fileFunction(self.excelFilePath,keywords)
            if(total == -1):
                self.totalRecordsFound.set("No Description Column Found")
            else:    
                self.totalRecordsFound.set(total)
        except PermissionError:
            self.totalRecordsFound.set("Please close the output file first.")




"""
    Now the only task left here is the

    PermissionError

"""
        

