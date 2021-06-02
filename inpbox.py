from tkinter import *
from tkinter.ttk import *
from tkinter.scrolledtext import ScrolledText
from tkinter.filedialog import askopenfilename,askdirectory
import tkinter.simpledialog as simpledialog
from tkinter import messagebox

import os
import logging
logger = logging.getLogger(f"MailMerge.{os.path.basename(__file__)}")

class InpBox:
    def __init__(self,title,LBox):
        self.title = title
        self.LBox = LBox
        
        inp = Toplevel()
        inp.wm_title(f"Add {title}")
        inp.resizable(False,False)
        inp.grab_set()
        inp.focus()

        self.root = inp

        if self.title == "Sender":
            self.add_sender_inp()
            self.add_data_inp(self.title + " email")
        else:
            self.add_data_inp(self.title)
        self.add_ok_btn()
        self.add_cancel_btn()
        
        self.root.mainloop()

    def add_sender_inp(self):
        self.sender_lbl = Label(self.root,text = f"Sender Name")
        self.sender_lbl.grid(row = 0,column = 0,padx = 10,pady = 10,columnspan = 2)
        self.sender_txt = Entry(self.root,font = "Consolas 12",width = "35")
        self.sender_txt.grid(row = 0,column = 3,padx =10,pady = 10,columnspan = 7)

    def add_data_inp(self,title):
        self.inp_lbl = Label(self.root,text = f"Enter {title}:")
        self.inp_lbl.grid(row = 2,column = 0,padx = 10,pady = 10,columnspan = 2)
        self.inp_txt = Entry(self.root,font = "Consolas 12",width = "35")
        self.inp_txt.grid(row = 2,column = 3,padx =10,pady = 10,columnspan = 7)

    def add_ok_btn(self):
        self.ok_btn = Button(self.root,text = "Ok",command = lambda:self.add())
        self.ok_btn.grid(row = 4,ipadx = 4, padx = 10,ipady = 3, pady = 10, column = 14, columnspan = 6,sticky = N+S+E+W)

    def add_cancel_btn(self):
        self.cancel_btn = Button(self.root,text = "Cancel",command = lambda:self.cancel())
        self.cancel_btn.grid(row = 4,ipadx = 4, padx = 10,ipady = 3, pady = 10, column = 21, columnspan = 6,sticky = N+S+E+W)

    def add(self):
        result = ""
        if "Sender" in self.title:
            if len(self.sender_txt.get()) == 0:
                messagebox.showerror("showerror","Enter Sender Name")
                self.root.mainloop()
                #logging.error("Sender Name Empty")
            else:
                result += self.sender_txt.get().upper()
        if len(self.inp_txt.get()) > 0:
            if "Sender" in self.title:
                result += f"<{self.inp_txt.get()}>"
            else:
                result += self.inp_txt.get()
            self.LBox.lbox.insert(END, result)
            self.cancel()
        else:
            self.cancel()
                
    def cancel(self):
        self.root.destroy()
        self.LBox.focus()
