from tkinter import *

import sys
import os
import glob
from pathlib import Path
from datetime import datetime
from string import Template
import json

import traceback
import logging

import ctypes
try: # >= win 8.1
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception as e:
    pass
try: # < win 8.1
    ctypes.windll.user32.SetProcessDPIAware()
except: # win 8.0 or less
    pass

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
        self.setup_logger()
        try:
            self.readJSON()
        except Exception:
            self.createJSON()
            pass

    def readJSON(self):
        with open("configuration.json","r") as fp:
            self.configData = json.load(fp)
            fp.close()

    def createJSON(self):
        dataMap = {
            "sender_name" : "",
            "sender_email" : "",
            "sender_designation" : "",
            "sender_department" : "",
            "org_name" : "",
            "org_group" : "",
            "org_addr" : "",
            "org_website" : "",
            "org_logo" : ""
        }
        self.dumpJSON(dataMap)

    def dumpJSON(self,dataMap):
        with open("configuration.json","w") as fp:
            json.dump(dataMap,fp)
        self.readJSON()

    def setup_logger(self):
        logging.basicConfig(level=logging.DEBUG, filename=self.LOGFILE, filemode="a+",format="%(asctime)-10s :: %(levelname)-8s --> %(message)s")
        logging.getLogger().addHandler(logging.StreamHandler(sys.stdout))
        self.logger = logging.getLogger(f"MailMerge.{os.path.basename(__file__)}")
    
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
        self.logger.info("Cleaning up Merged docs")
        doc_files_path = os.path.join(self.DOCS_PATH,"*")
        files = glob.glob(doc_files_path)
        for file in files:
            try:
                with open(file) as fp:
                    fp.close()
                os.remove(file)
            except Exception as e:
                self.logger.exception(e)

    def get_display_size(self):
        root = Tk()
        root.update_idletasks()
        root.attributes('-fullscreen', True)
        root.state('iconic')
        height = root.winfo_screenheight()
        width = root.winfo_screenwidth()
        root.destroy()
        root.quit()
        return height, width

    def setup_signature(self):
        #print(self.configData)
        sender_details = self.setup_sender_details().substitute(**self.configData)
        org_details = self.setup_organisation().substitute(**self.configData)

        org_img = '<img src="cid:image1" width="80" height="60">'

        self.signature = f"""Regards,\n{sender_details}\n\n<table style="font-size:80%;"><tr><td>{org_img}</td><td>{org_details}</td></tr></table>"""
        self.signature = self.signature.replace("\n","<br/>")
        return self.signature

    def setup_sender_details(self):
        self.sender_details = "$sender_name\n$sender_designation\n$sender_department"
        self.sender_details = self.sender_details.replace("\n","<br/>")

        self.sender_details = Template(self.sender_details)

        return self.sender_details

    def setup_organisation(self):
        self.org_details = '$org_name\n$org_group\n$org_addr\nWebsite: <a href="$org_website">$org_website</a>'
        self.org_details = self.org_details.replace("\n","<br/>")

        self.org_details = Template(self.org_details)

        return self.org_details
