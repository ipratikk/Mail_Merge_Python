import os
from mailmerge import MailMerge

class ReadTemplate:
    def __init__(self,filename):
        self.FILE = MailMerge(filename)

    def merge_fields(self):
        return self.FILE.get_merge_fields()

    def populate_fields(self,fields):
        self.FILENAME = f"{fields['Email']}.docx"
        self.FILE.merge(**fields)

    def save(self,dirpath):
        filepath = os.path.join(dirpath,self.FILENAME)
        self.FILE.write(filepath)
