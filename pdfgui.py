from tkinter import *
from tkinter.ttk import *
from tkinter.scrolledtext import ScrolledText
from tkinter.filedialog import askopenfilename,askdirectory
from tkinter import messagebox

from setup import Setup
from tkPDFViewer import tkPDFViewer as pdf
from mailgui import MailGUI

import os
import logging
logger = logging.getLogger(f"MailMerge.{os.path.basename(__file__)}")


class PDF_Preview:
    def __init__(self,data,width=1600,height=450):
        self.screen_height = height
        self.screen_width = width

        self.height = self.screen_height - 100
        self.width = self.screen_width // 2

        self.setup = Setup()
        self.data = data
        self.idx_str = StringVar()

        self.idx = [0]
        self.pdf_files = data['Email'].to_list()
        self.total = len(self.pdf_files)
        
        logger.info("Initialising PDF Preview GUI")
        top = Toplevel()
        top.wm_title("Mail Merge Preview")
        #top.geometry(f"{self.width}x{self.height}")
        top.maxsize(self.width,self.height)
        top.resizable(True,True)
        top.grab_set()
        top.lift()

        self.root = top

        self.add_left_btn()
        self.add_right_btn()
        self.add_idx_lbl()
        self.add_cancel_btn()
        self.add_send_btn()

        try:
            pdf_path = os.path.join(self.setup.PDF_PATH,f"{self.pdf_files[0]}.pdf")
            self.open_file(pdf_path)
        except Exception:
            pass
    
        
        self.root.mainloop()

    def open_file(self,pdf_path):
        v1 = pdf.ShowPdf()
        try:
            v1.img_object_li.pop(0)
        except Exception:
            pass
        v2 = v1.pdf_view(self.root, pdf_location = pdf_path ,width = 75, height = 40, bar = False)
        v2.grid(row = 7, padx = 10, pady = 10, column = 1, columnspan = 60,sticky = N+S+E+W)
        self.idx_str.set(f"Preview : {self.idx[0]+1}/{self.total}")

    def prev_file(self,idx):
        idx[0] = max(idx[0]-1,0)
        self.idx = idx
        self.right_btn['state'] = NORMAL
        if idx[0] == 0:
            self.left_btn['state']=DISABLED
        else:
            self.left_btn['state']=NORMAL
        pdf_path = os.path.join(self.setup.PDF_PATH, f"{self.pdf_files[idx[0]]}.pdf")
        print("Prev",idx[0],pdf_path)
        self.open_file(pdf_path)

    def next_file(self,idx):
        idx[0] = min(idx[0]+1,len(self.pdf_files)-1)
        self.idx = idx
        self.left_btn['state'] = NORMAL
        if idx[0] == len(self.pdf_files)-1:
            self.right_btn['state']=DISABLED
        else:
            self.right_btn['state']=NORMAL
        pdf_path = os.path.join(self.setup.PDF_PATH, f"{self.pdf_files[idx[0]]}.pdf")
        print("Next",idx[0],pdf_path)
        self.open_file(pdf_path)

    def add_left_btn(self):
        self.left_btn = Button(self.root,text = "<",width=10,command = lambda:self.prev_file(self.idx))
        self.left_btn.grid(row = 3, padx = 10, pady = 10, column = 1, columnspan = 1,sticky = N+S+E+W)

    def add_idx_lbl(self):
        self.idx_lbl = Label(self.root,font = "Consolas 12",textvariable = self.idx_str,width = "30",state="readonly")
        self.idx_lbl.grid(row = 3, padx = 10, pady = 10, column = 15, columnspan = 17,sticky = N+S+E+W)

    def add_right_btn(self):
        self.right_btn = Button(self.root,text = ">",width=10,command = lambda:self.next_file(self.idx))
        self.right_btn.grid(row = 3, padx = 10, pady = 10, column = 60, columnspan = 1,sticky = N+S+E+W)

    def add_cancel_btn(self):
        self.cancel_btn = Button(self.root,text = "Cancel",width=10,command = self.root.destroy)
        self.cancel_btn.grid(row = 1, padx = 10, pady = 10, column = 57, columnspan = 1,sticky = N+S+E+W)

    def add_send_btn(self):
        self.send_btn = Button(self.root,text = "Send Emails",width=10,command = self.send_emails)
        self.send_btn.grid(row = 1, padx = 10, pady = 10, column = 60, columnspan = 1,sticky = N+S+E+W)

    def send_emails(self):
        logger.info("Closing PDF Preview")
        self.root.destroy()
        logger.info("Sending Emails")
        MailGUI(self.data)
    
