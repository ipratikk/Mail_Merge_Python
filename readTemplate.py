import os
from mailmerge import MailMerge

import logging
logger = logging.getLogger(f"MailMerge.{os.path.basename(__file__)}")

class ReadTemplate:
    def __init__(self,filename):
        self.FILE = MailMerge(filename)

    def merge_fields(self):
        logger.info("Geting merge fields")
        return self.FILE.get_merge_fields()

    def populate_fields(self,fields):
        self.FILENAME = f"{fields['Email']}.docx"
        self.FILE.merge(**fields)
        logger.info("Populating merge fields")

    def save(self,dirpath):
        filepath = os.path.join(dirpath,self.FILENAME)
        self.FILE.write(filepath)
        logger.info(f"Saving file {self.FILENAME}")
