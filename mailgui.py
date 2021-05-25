from tkinter import *
from tkinter.ttk import *
from tkinter.scrolledtext import ScrolledText
from tkinter.filedialog import askopenfilename,askdirectory
import tkinter.simpledialog as simpledialog
from tkinter import messagebox

from lbox import LBox
from config import Config_data

class MailGUI:
    def __init__(self,app_name,email_list):
        self.email_list = email_list

        root = Tk()
        root.wm_title(app_name)
        root.minsize(500,250)

        self.root = root
        self.add_server_picker()
        self.add_sender_email()
        self.add_receipent_email()
        self.add_cc_emails()
        self.add_send_btn()
        self.add_copyright_lbl()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.resizable(False,False)

        self.root.mainloop()

    def on_closing(self):
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            self.root.destroy()
            self.root.quit()

    def add_server_picker(self):
        self.server_lbl = Label(self.root,text = "Email Server")
        self.server = Combobox(self.root,width="5",state = "readonly")
        self.server['values'] = ('Gmail','Outlook')
        self.server_lbl.grid(row = 1, padx = 10 , pady = 10, column = 0, columnspan = 2,sticky = N+S+E+W)
        self.server.grid(row = 1, padx = 10 , pady = 10, column = 3, columnspan = 1,sticky = N+S+E+W)
        self.server.current(1)

    def add_sender_email(self):
        self.sender_email_str = StringVar()
        self.sender_email_lbl = Label(self.root,text = "Sender Email")
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

    def add_send_btn(self):
        self.send_mail = Button(self.root,text = "Send Email",command = lambda:self.run_script(self.server,self.sender_email,self.receipent,self.cc))
        self.send_mail.grid(row = 21,ipadx = 4, padx = 10,ipady = 3, pady = 10, column = 20, columnspan = 6,sticky = N+S+E+W)

    def add_copyright_lbl(self):
        copyright = "© 2020 Ernst & Young LLP, Published in India"
        self.cp_lbl = Label(self.root,text = copyright,)
        self.cp_lbl.grid(row = 24, padx = 10 , pady = 10, column = 4, columnspan = 11,sticky = N+S+E+W)

    def run_script(self):
        print("Running")
        
