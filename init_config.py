from tkinter import *
from tkinter.ttk import *
from tkinter.scrolledtext import ScrolledText
from tkinter.filedialog import askopenfilename,askdirectory
from tkinter import messagebox

from setup import Setup

import threading
import os
import logging
logger = logging.getLogger(f"MailMerge.{os.path.basename(__file__)}")

class InitSetup:
	def __init__(self):
		self.setup = Setup()
		self.data = self.setup.configData
		root = Toplevel()
		root.wm_title("Mail Merge Utility")
		root.minsize(500,200)

		self.root = root

		self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
		self.root.resizable(False,False)

		self.add_seperator("Sender",row=1)
		self.sender_name_str = StringVar()
		self.sender_email_str = StringVar()
		self.sender_designation_str = StringVar()
		self.sender_department_str = StringVar()

		self.add_record("Name","sender_name",self.sender_name_str,row=2)
		self.add_record("Email","sender_email",self.sender_email_str,row=3)
		self.add_record("Designation","sender_designation",self.sender_designation_str,row=4)
		self.add_record("Department","sender_department",self.sender_department_str,row=5)

		self.add_seperator("Organisation",row=7)
		self.org_name_str = StringVar()
		self.org_group_str = StringVar()
		self.org_addr_str = StringVar()
		self.org_website_str = StringVar()
		self.org_logo_str = StringVar()
		
		self.add_record("Name","org_name",self.org_name_str,row=8)
		self.add_record("Group","org_group",self.org_group_str,row=9)
		self.add_record("Address","org_addr",self.org_addr_str,row=10,longText=True)
		self.add_record("Website","org_website",self.org_website_str,row=11)
		self.add_org_img_inp("org_logo",self.org_logo_str,row=12)
		
		self.add_generate_btn(row=14)
		logger.info("Initialising copyright Label")
		self.add_copyright_lbl(row=15)


		self.root.mainloop()

	def on_closing(self):
		if messagebox.askokcancel("Quit", "Do you want to quit?"):
			self.root.destroy()
			self.root.quit()
			logger.info("Setting Up initial Configuration")

	def add_seperator(self,title,row=1):
		self.sep = Separator(self.root, orient='horizontal')
		self.sep.grid(row = row,column = 0,columnspan = 40,padx=10,pady=0,sticky="ew")
		self.inp_lbl = Label(self.root,text = f"{title}",font = ("Arial",14))
		self.inp_lbl.grid(row = row,column = 1,padx = 10,pady = 10,columnspan = 2)

	def add_record(self,title,key,inp_str,row=1,longText=False):
		self.inp_str.set(self.data[key])
		inp_lbl = Label(self.root,text = f"{title}")
		inp_lbl.grid(row = row,column = 0,padx = 10,pady = 10,columnspan = 2)
		if longText:
			inp_txt = ScrolledText(self.root,font = "Consolas 10",width = "43",height=2,wrap=WORD)
		else:
			inp_txt = Entry(self.root,font = "Consolas 12",width = "35",textvariable=inp_str)
		inp_txt.grid(row = row,column = 3,padx =10,pady = 10,columnspan = 7)
		
	
	def add_org_img_inp(self,inp_str,row=1):
		self.org_img_str = StringVar()
		self.org_img = Label(self.root,text = "Organization Icon")
		self.org_img.grid(row = row, padx = 10 , pady = 10, column = 0, columnspan = 2,sticky = N+S+E+W)
		self.org_img_dir = Entry(self.root,font = "Consolas 12",textvariable = inp_str,width = "30",state="readonly")
		self.org_img_dir.grid(row = row, column = 3,padx = 10, pady = 10,columnspan = 15,sticky = N+S+E+W)
		self.org_img_browse = Button(self.root,text = "Browse",command = lambda:self.get_logo(inp_str))
		self.org_img_browse.grid(row = row+1, padx = 10, pady = 10, column = 9, columnspan = 1,sticky = N+S+E+W)

	def add_generate_btn(self,row=1):
		self.gen = Button(self.root,text = "Configure",command = lambda:self.run_script)
		self.gen.grid(row = row, padx = 10, pady = 10, column = 2, columnspan = 10,sticky = N+S+E+W)

	def add_copyright_lbl(self,row=1):
		copyright = "© 2021 Pratik Goel, Published in India"
		self.cp_lbl = Label(self.root,text = copyright)
		self.cp_lbl.grid(row = row, padx = 10 , pady = 10, column = 2, columnspan = 12,sticky = N+S+E+W)

	def get_logo(self):
		filename = askopenfilename(filetypes = [('PNG file','*.png')])
		self.org_img_str.set(filename)

	def run_script(self):
		dataMap = {
			"sender_name" : self.sender_name_str,
			"sender_email" : self.sender_email_str,
			"sender_designation" : self.sender_designation_str,
			"sender_department" : self.sender_department_str,
			"org_name" : self.org_name_str,
			"org_group" : self.org_group_str,
			"org_addr" : self.org_addr_str,
			"org_website" : self.org_website_str,
			"org_logo" : self.org_img_str
		}

if __name__ == "__main__":
	InitSetup()