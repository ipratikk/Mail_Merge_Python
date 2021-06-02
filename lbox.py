from tkinter import *
from tkinter.ttk import *
from tkinter.scrolledtext import ScrolledText
from tkinter.filedialog import askopenfilename,askdirectory
import tkinter.simpledialog as simpledialog
from tkinter import messagebox

from inpbox import InpBox
from config import Config_data

import os
import logging
logger = logging.getLogger(f"MailMerge.{os.path.basename(__file__)}")

class LBox(object):
    def __init__(self,title,strVar,items=None,select=MULTIPLE,isEditable=True):
        self.config = Config_data()
        self.title = title
        self.items = self.config.config_data[title]["Items"]
        if items:
            self.items = items
        self.select = select
        self.strVar = strVar
        self.isEditable = isEditable
        self.setup_frame(isEditable)
        
    def setup_frame(self,isEditable):
        top = Toplevel()
        top.wm_title(f"Select the {self.title}")
        top.minsize(300,250)
        top.resizable(False,False)
        top.grab_set()
        top.focus()
        
        self.root = top

        self.add_top_frame()
        self.add_bottom_frame()
        self.add_left_frame(pady = 15,padx = (20,0),side = LEFT)
        if isEditable:
            self.add_right_frame(pady = 50,padx = 10)
        self.add_yscrollbar(side = RIGHT, fill = Y)
        self.set_select_method(self.select)
        self.add_lbox()
        self.config_yscrollbar()
        self.add_ok_btn(fill = BOTH, expand = 1, padx = 40,pady = 10)

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.resizable(False,False)

        self.root.mainloop()

    def on_closing(self):
        self.update_config()
        
    def add_top_frame(self,**args):
        self.top_frame = Frame(self.root)
        self.top_frame.pack(args)

    def add_bottom_frame(self,**args):
        self.bottom_frame = Frame(self.root)
        self.bottom_frame.pack(args)

    def add_left_frame(self,**args):
        self.left_frame = Frame(self.top_frame)
        self.left_frame.pack(args)

    def add_right_frame(self,**args):
        self.right_frame = Frame(self.top_frame)
        add_btn = Button(self.right_frame,text = "Add",command = self.insert_item)
        remove_btn = Button(self.right_frame,text = "Remove",command = self.delete_item)

        add_btn.pack(pady = 5)
        remove_btn.pack(pady = 5)

        self.right_frame.pack(args)

    def add_yscrollbar(self,**args):
        self.yscrollbar = Scrollbar(self.left_frame,orient= VERTICAL)
        self.yscrollbar.pack(args)

    def set_select_method(self,select):
        self.select = select

    def add_lbox(self,**args):
        self.lbox = Listbox(self.left_frame,width = 80, selectmode = self.select, yscrollcommand = self.yscrollbar.set)
        self.lbox.insert(END,*self.items)
        self.lbox.pack()

    def config_yscrollbar(self):
        self.yscrollbar.config(command = self.lbox.yview)

    def add_ok_btn(self,**args):
        self.ok_btn = Button(self.bottom_frame,text = "OK",width = 50,command=self.update_config)
        self.ok_btn.pack(args)

    def insert_item(self):
        InpBox(self.title,self)
        self.focus()

    def delete_item(self):
        result = messagebox.askquestion("Remove", "Are You Sure?", icon='warning')
        if result == 'yes':
            for idx in self.lbox.curselection():
                self.lbox.delete(idx)

    def focus(self):
        self.root.lift()
        self.root.grab_set()
        self.root.focus()

    def update_config(self):
        if not self.isEditable:
            self.root.destroy()
            return
        items = list(self.lbox.get(0,END))
        selected = [self.lbox.get(idx) for idx in self.lbox.curselection()]
        if self.title == "Sender":
            self.config.update_sender(items)
        elif self.title == "To":
            selected = items
        elif self.title == "CC":
            self.config.update_cc(items)
        self.strVar.set(",".join(selected))
        self.root.destroy()
    
