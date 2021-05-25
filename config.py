import os
import json
from pathlib import Path
from collections import defaultdict


class Config_data:

    def __init__(self,email_list=None):
        self.email_list = email_list
        self.app_data_path = os.getenv("LOCALAPPDATA")
        self.temp_data_path = os.path.join(self.app_data_path,"Mail_Merge")
        self.config_file = os.path.join(self.temp_data_path,"configuration.json")
        self.config_data = self.read_json()
        self.sender_data = self.load_sender()
        self.to_data = self.load_to()
        self.cc_data = self.load_cc()

    def load_sender(self):
        data = defaultdict(str)
        if "Sender" in self.config_data:
            return self.config_data['Sender']
        return data
    
    def load_to(self):
        data = {'Desc':'Receipent Email','Items':self.email_list}
        return data

    def load_cc(self):
        data = defaultdict(str)
        if "CC" in self.config_data:
            return self.config_data['CC']
        return data

    def update_sender(self,items):
        self.sender_data['Desc'] = "Sender Email"
        self.sender_data['Items'] = items
        self.config_data['Sender'] = self.sender_data
        self.dump_json(self.config_file,self.config_data)
        

    def update_to(self,items):
        self.to_data['Desc'] = "Receipent Email"
        self.to_data['Items'] = self.email_list
        self.config_data['To'] = self.to_data
        self.dump_json(self.config_file,self.config_data)

    def update_cc(self,items):
        self.cc_data['Desc'] = "CC Email(s)"
        self.cc_data['Items'] = items
        self.config_data['CC'] = self.cc_data
        self.dump_json(self.config_file,self.config_data)

    def read_json(self):
        data = None
        if not self.check_exists(self.config_file):
            self.create_tree(self.config_file)
        with open(self.config_file,"r") as fp:
            data = json.load(fp)
            data['To'] = self.load_to()
            fp.close()
        return data

    def dump_json(self,config_file,data):
        with open(config_file,"w") as fp:
            json.dump(data,fp)
            fp.close()

    def check_exists(self,config_file):
        if os.path.exists(config_file):
            return True
        return False

    def create_tree(self,config_file):
        Path(self.temp_data_path).mkdir(parents=True, exist_ok=True)
        data = defaultdict(list)
        data['Sender'] = {'Desc':'Sender Email','Items':[]}
        data['To'] = {'Desc':'Receipent Email','Items':[]}
        data['CC'] = {'Desc':'CC Email(s)','Items':[]}
        self.dump_json(config_file,data)
