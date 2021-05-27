import sys
import os
import glob
from pathlib import Path
from datetime import datetime
from string import Template
import json

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
        try:
            with open("configuration.json","r") as fp:
                self.configData = json.load(fp)
                fp.close()
        except Exception:
            pass
    
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

    def setup_signature(self):
        print(self.configData)
        sender_details = self.setup_sender_details().substitute(**self.configData)
        org_details = self.setup_organisation().substitute(**self.configData)

        org_img = '<img src="cid:image1" width="80" height="60">'

        self.signature = f"""Regards,\n{sender_details}\n\n<table style="font-size:60%;"><tr><td>{org_img}</td><td>{org_details}</td></tr></table>"""
        self.signature = self.signature.replace("\n","<br/>")
        return self.signature

    def setup_sender_details(self):
        #self.sender_name = self.configData['sender_name']
        #self.sender_designation = "President"
        #self.sender_department = "Codechef UEMK Chapter"

        self.sender_details = "$sender_name\n$sender_designation\n$sender_department"
        self.sender_details = self.sender_details.replace("\n","<br/>")

        self.sender_details = Template(self.sender_details)

        return self.sender_details

    def setup_organisation(self):
        #self.org_name = "University Of Engineering & Management, Kolakta"
        #self.org_group = "(IEM-UEM Group)"
        #self.org_addr = "University Area, Plot No. III - B/5\nMain Arterial Road, New Town\nAction Area - III, Kolkata 700 160"
        #self.org_website = "http://www.uem.edu.in/"

        self.org_details = '$org_name\n$org_group\n$org_addr\nWebsite: <a href="$org_website">$org_website</a>'
        self.org_details = self.org_details.replace("\n","<br/>")

        self.org_details = Template(self.org_details)

        return self.org_details
