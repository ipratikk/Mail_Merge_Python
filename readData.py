import pandas as pd
import setup

class ReadExcel:

    def __init__(self,filename,sheetName=""):
        self.FILENAME = filename
        self.SHEETNAME = sheetName
    
    def read_data(self):
        if self.SHEETNAME != "":
            data = pd.read_excel(self.FILENAME,sheet_name = self.SHEETNAME, dtype=str)
        else:
            data = pd.read_excel(self.FILENAME, dtype=str)
        
        self.DATA = self.clean_data(data)
        return self.DATA
    
    def clean_data(self,data):
        data = data.dropna(axis = 0, how = "all")
        return data
