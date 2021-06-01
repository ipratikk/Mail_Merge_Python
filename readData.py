import pandas as pd
import setup

import traceback
import os
import logging
logger = logging.getLogger(f"MailMerge.{os.path.basename(__file__)}")

class ReadExcel:

    def __init__(self,filename,sheetName=""):
        self.FILENAME = filename
        self.SHEETNAME = sheetName
    
    def read_data(self):
        logger.info("Reading Data from Excel")
        try:
            if self.SHEETNAME != "":
                data = pd.read_excel(self.FILENAME,sheet_name = self.SHEETNAME, dtype=str)
            else:
                data = pd.read_excel(self.FILENAME, dtype=str)
        except Exception as e:
            traceback.print_exc()
            logger.exception(e)
        
        self.DATA = self.clean_data(data)
        return self.DATA
    
    def clean_data(self,data):
        logger.info("Cleaning Excel data")
        data = data.dropna(axis = 0, how = "all")
        return data
