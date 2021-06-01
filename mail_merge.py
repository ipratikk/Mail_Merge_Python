from datetime import datetime
from setup import Setup
from readData import ReadExcel
from readTemplate import ReadTemplate
from docx2pdf import convert
import threading
import traceback
import os

import logging
logger = logging.getLogger(f"MailMerge.{os.path.basename(__file__)}")

class MailMerge:
    def __init__(self,template_file="",excel_file=""):
        self.TEMPLATE_FILE = template_file
        self.EXCEL_FILE = excel_file

    def init_setup(self):
        self.setup = Setup()

    def read_template(self,template=None):
        if template == None:
            template = self.TEMPLATE_FILE
        obj = ReadTemplate(template)
        self.TEMPLATE_DATA = obj
        return obj

    def read_excel(self,excel=None):
        if excel == None:
            excel = self.EXCEL_FILE
        obj = ReadExcel(excel)
        self.EXCEL_DATA = obj
        return obj

    def start_merge(self,progress_lbl=None,progress=None):
        self.init_setup()
        self.read_template()
        self.read_excel()
        data = self.EXCEL_DATA.read_data()
        total = len(data.index)
        cnt = 100 // total
        idx = 1
        for fields in data.to_dict(orient='records'):
            self.merge_doc(fields)
            if progress != None:
                progress['value'] += cnt
                progress_lbl.set(f"Merging ({idx} / {total})")
                logger.info(f"Merging ({idx} / {total})")
        convert(self.setup.DOCS_PATH,self.setup.PDF_PATH)
        self.setup.cleanup()

    def merge_doc(self,fields):
        if 'Date' not in fields:
            fields['Date'] = datetime.now().strftime("%d %B, %Y")
        doc = self.read_template()
        doc.populate_fields(fields)
        doc.save(self.setup.DOCS_PATH)

    def convert_pdf(self):
        convert(self.setup.DOCS_PATH,self.setup.PDF_PATH)
        self.setup.cleanup()
            
