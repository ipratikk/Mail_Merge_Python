import threading
import smtplib 
from tkinter import *
from tkinter.ttk import *
from tkinter.scrolledtext import ScrolledText
from tkinter.filedialog import askopenfilename,askdirectory
import tkinter.simpledialog as simpledialog
from tkinter import messagebox

from lbox import LBox
from config import Config_data
from imap_email import Send_Mail


import os
import logging
logger = logging.getLogger(f"MailMerge.{os.path.basename(__file__)}")

class MailGUI:
    def __init__(self,data):
        self.data = data
        self.email_list = self.data['Email'].to_list()

        root = Toplevel()
        root.wm_title("Email Client")
        root.minsize(500,250)
        root.grab_set()
        root.lift()

        self.root = root
        self.add_server_picker()
        self.add_sender_email()
        self.add_receipent_email()
        self.add_cc_emails()
        self.add_subject_inp()
        self.add_body_inp()
        self.add_send_btn()
        self.add_copyright_lbl()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.resizable(False,False)

        self.root.mainloop()

    def on_closing(self):
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            self.root.destroy()

    def show_alert(self,code,message):
        if code == "showinfo":
            messagebox.showinfo(code,message)
            self.root.lift()
            logger.info(f"{message}")
            return
        messagebox.showerror(code,message)
        logger.error(f"{message}")

    def add_server_picker(self):
        self.server_lbl = Label(self.root,text = "Email Server")
        self.server = Combobox(self.root,width="5",state = "readonly")
        self.server['values'] = ('Gmail','Outlook')
        self.server_lbl.grid(row = 1, padx = 10 , pady = 10, column = 0, columnspan = 2,sticky = N+S+E+W)
        self.server.grid(row = 1, padx = 10 , pady = 10, column = 3, columnspan = 1,sticky = N+S+E+W)
        self.server.current(1)

    def add_sender_email(self):
        self.sender_email_str = StringVar()
        self.sender_email_lbl = Label(self.root,text = "Sender")
        self.sender_email_lbl.grid(row = 1, padx = 10 , pady = 10, column = 4, columnspan = 4,sticky = N+S+E+W)
        self.sender_email = Entry(self.root,font = "Consolas 10",textvariable = self.sender_email_str,state = "readonly")
        self.sender_email.grid(row = 1, padx = 10 , pady = 10, column = 8, columnspan = 12,sticky = N+S+E+W)

        self.sender_email_select = Button(self.root,text = "Select",command = lambda:LBox("Sender",self.sender_email_str,select=SINGLE))
        self.sender_email_select.grid(row = 1, padx = 10, pady = 10, column = 20, columnspan = 3,sticky = N+S+E+W)

    def add_receipent_email(self):
        self.receipent_str = StringVar()
        self.receipent_str.set(", ".join(self.email_list))
        self.receipent_lbl = Label(self.root,text = "Receipent Email")
        self.receipent_lbl.grid(row = 3, padx = 10 , pady = 10, column = 0, columnspan = 2,sticky = N+S+E+W)
        self.receipent = Entry(self.root,font = "Consolas 10",textvariable = self.receipent_str,width = "50",state = "readonly")
        self.receipent.grid(row = 3, column = 3,padx = 10, pady = 10,columnspan = 15,sticky = N+S+E+W)
        self.receipent_select = Button(self.root,text = "Select",command = lambda:LBox("To",self.receipent_str,items=self.email_list,isEditable=False))
        self.receipent_select.grid(row = 3, padx = 10, pady = 10, column = 20, columnspan = 3,sticky = N+S+E+W)

    def add_cc_emails(self):
        self.cc_str = StringVar()
        self.cc_lbl = Label(self.root,text = "CC")
        self.cc_lbl.grid(row = 5, padx = 10 , pady = 10, column = 0, columnspan = 2,sticky = N+S+E+W)
        self.cc = Entry(self.root,font = "Consolas 10",textvariable = self.cc_str,width = "50",state = "readonly")
        self.cc.grid(row = 5, column = 3,padx = 10, pady = 10,columnspan = 15,sticky = N+S+E+W)
        self.cc_select = Button(self.root,text = "Select",command = lambda:LBox("CC",self.cc_str))
        self.cc_select.grid(row = 5, padx = 10, pady = 10, column = 20, columnspan = 3,sticky = N+S+E+W)

    def add_subject_inp(self):
        self.subject_str = StringVar()
        self.subject_lbl = Label(self.root,text = "Subject")
        self.subject_lbl.grid(row = 7, padx = 10 , pady = 10, column = 0, columnspan = 2,sticky = N+S+E+W)
        self.subject = Entry(self.root,font = "Consolas 10",textvariable = self.subject_str,width = "50")
        self.subject.grid(row = 7, column = 3,padx = 10, pady = 10,columnspan = 15,sticky = N+S+E+W)

    def add_body_inp(self):
        self.body_lbl = Label(self.root,text = "Body")
        self.body_lbl.grid(row = 9, padx = 10 , pady = 10, column = 0, columnspan = 2,sticky = N+S+E+W)
        self.body = ScrolledText(self.root,font = "Consolas 10",width = "50",height=6,wrap=WORD)
        self.body.grid(row = 9, column = 3,padx = 10, pady = 10,columnspan = 15,sticky = N+S+E+W)

    def add_send_btn(self):
        self.send_mail = Button(self.root,text = "Send Email",command = self.run_script)
        self.send_mail.grid(row = 21,ipadx = 4, padx = 10,ipady = 3, pady = 10, column = 20, columnspan = 6,sticky = N+S+E+W)

    def add_copyright_lbl(self):
        copyright = "© 2021 Pratik Goel, Published in India"
        self.cp_lbl = Label(self.root,text = copyright,)
        self.cp_lbl.grid(row = 24, padx = 10 , pady = 10, column = 4, columnspan = 11,sticky = N+S+E+W)

    def run_script(self):
        #Tk().withdraw()
        self.pwd= simpledialog.askstring("Password", "Enter password:", show='*')
        self.root.lift()
        
        thread1 = threading.Thread(target = self.send_emails)
        thread1.start()
    
    def send_emails(self):
        self.var = StringVar()

        self.status_lbl = Label(self.root,textvariable = self.var)
        self.status_lbl.grid(row = 13, padx = 10 , pady = 10, column = 1, columnspan = 20,sticky = N+S+E+W)
        
        self.progress = Progressbar(self.root,orient = HORIZONTAL, length = 100, mode = 'determinate')
        self.progress.grid(row = 15, padx = 10, pady = 10, column = 1, columnspan = 60,sticky = N+S+E+W)
        self.progress['value'] = 0
        
        obj = Send_Mail(self.server,self.sender_email_str,self.receipent_str,self.cc_str,self.subject_str,self.body,self.data,status_str=self.var,progressbar=self.progress)
        #pwd = simpledialog.askstring(title="Authenticate Email",prompt=f"Enter password:")
        try:
            obj.login(password=self.pwd)
        except smtplib.SMTPAuthenticationError:
            self.show_alert("showerror","Authentication Error, Check credentials")
            self.var.set("Authentication Error!")
            return

        sent = obj.send_emails()
        if sent:
            self.show_alert("showinfo","Sent all emails")
        else:
            self.show_alert("showerror","Some Error Occured, Checks logs for more!")
