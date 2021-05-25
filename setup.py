import sys
import os
import glob
from pathlib import Path
from datetime import datetime

class Setup:
    
    def __init__(self):
        self.DATA_PATH = ""
        self.setup()
        print(self.DATA_PATH)
        print(self.LOGFILE)

    def setup(self):
        self.setup_path()
        self.LOGFILE = os.path.join(self.DATA_PATH,f'Mail_Merge Log {datetime.now().strftime("%Y-%m-%d_%H-%M")}.log')
        self.DOCS_PATH = os.path.join(self.DATA_PATH,"Docs")
        self.PDF_PATH = os.path.join(self.DATA_PATH,"Pdf")
        if not os.path.exists(self.LOGFILE):
            Path(self.DATA_PATH).mkdir(parents=True, exist_ok=True)
            Path(self.DOCS_PATH).mkdir(parents=True, exist_ok=True)
            Path(self.PDF_PATH).mkdir(parents=True, exist_ok=True)
    
    def setup_path(self):
        PLATFORM = sys.platform
        if PLATFORM == "darwin":
            desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
            self.DATA_PATH = os.path.join(desktop_path,'Mail_Merge')
        elif PLATFORM == "win32":
            app_data_path = os.getenv("LOCALAPPDATA")
            temp_data_path = os.path.join(app_data_path,"Mail_Merge")
            self.DATA_PATH = temp_data_path

    def cleanup(self):
        doc_files_path = os.path.join(self.DOCS_PATH,"*")
        files = glob.glob(doc_files_path)
        for file in files:
            os.remove(file)
