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
	def __init__(self,parent):
		self.parent = parent
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
		self.add_address_inp("Address","org_addr",row=10)
		self.add_record("Website","org_website",self.org_website_str,row=11)
		self.add_org_img_inp("org_logo",self.org_logo_str,row=12)
		
		self.add_configure_btn(row=14)
		logger.info("Initialising copyright Label")
		self.add_copyright_lbl(row=15)

		self.root.lift()
		self.root.mainloop()

	def on_closing(self):
		if messagebox.askokcancel("Quit", "Do you want to quit?"):
			self.lift_parent()
			self.root.destroy()
			logger.info("Setting Up initial Configuration")

	def lift_parent(self):
		self.parent.update()
		self.parent.deiconify()
		self.parent.lift()

	def show_alert(self,code,message):
		if code == "showinfo":
		    messagebox.showinfo(code,message)
		    self.root.lift()
		    logger.info(f"{message}")
		    return
		messagebox.showerror(code,message)
		logger.error(f"{message}")

	def add_seperator(self,title,row=1):
		self.sep = Separator(self.root, orient='horizontal')
		self.sep.grid(row = row,column = 0,columnspan = 40,padx=10,pady=0,sticky="ew")
		self.inp_lbl = Label(self.root,text = f"{title}",font = ("Arial",14))
		self.inp_lbl.grid(row = row,column = 1,padx = 10,pady = 10,columnspan = 2)

	def add_record(self,title,key,inp_str,row=1):
		inp_str.set(self.data[key])
		inp_lbl = Label(self.root,text = f"{title}")
		inp_lbl.grid(row = row,column = 0,padx = 10,pady = 10,columnspan = 2)
		inp_txt = Entry(self.root,font = "Consolas 12",width = "35",textvariable=inp_str)
		inp_txt.grid(row = row,column = 3,padx =10,pady = 10,columnspan = 7)

	def add_address_inp(self,title,key,row=1):
		self.addr_lbl = Label(self.root,text = f"{title}")
		self.addr_lbl.grid(row = row,column = 0,padx = 10,pady = 10,columnspan = 2)
		self.addr_txt = ScrolledText(self.root,font = "Consolas 10",width = "43",height=2,wrap=WORD)
		self.addr_txt.grid(row = row,column = 3,padx =10,pady = 10,columnspan = 7)
		addr = self.data[key]
		self.addr_txt.delete('1.0',END)
		self.addr_txt.insert('1.0',addr)
		
	
	def add_org_img_inp(self,key,inp_str,row=1):
		inp_str.set(self.data[key])
		self.org_img = Label(self.root,text = "Organization Icon")
		self.org_img.grid(row = row, padx = 10 , pady = 10, column = 0, columnspan = 2,sticky = N+S+E+W)
		self.org_img_dir = Entry(self.root,font = "Consolas 12",textvariable = inp_str,width = "30",state="readonly")
		self.org_img_dir.grid(row = row, column = 3,padx = 10, pady = 10,columnspan = 15,sticky = N+S+E+W)
		self.org_img_browse = Button(self.root,text = "Browse",command = lambda:self.get_logo(inp_str))
		self.org_img_browse.grid(row = row+1, padx = 10, pady = 10, column = 9, columnspan = 1,sticky = N+S+E+W)

	def add_configure_btn(self,row=1):
		self.gen = Button(self.root,text = "Configure",command = self.run_script)
		self.gen.grid(row = row, padx = 10, pady = 10, column = 2, columnspan = 10,sticky = N+S+E+W)

	def add_copyright_lbl(self,row=1):
		copyright = "© 2021 Pratik Goel, Published in India"
		self.cp_lbl = Label(self.root,text = copyright)
		self.cp_lbl.grid(row = row, padx = 10 , pady = 10, column = 2, columnspan = 12,sticky = N+S+E+W)

	def get_logo(self,inp_str):
		filename = askopenfilename(filetypes = [('PNG file','*.png')])
		inp_str.set(filename)

	def validate_fields(self,fields):
		for field in fields:
			inp_str,title = field
			if title == "Organization Address":
				if len(inp_str.get('1.0','end-1c')) == 0:
					self.show_alert("showerror",f"{title} cannot be empty")
					return False	
			elif len(inp_str.get()) == 0:
				self.show_alert("showerror",f"{title} cannot be empty")
				return False
		return True

	def run_script(self):
		fields = [
					(self.sender_name_str,"Sender Name"),
					(self.sender_email_str,"Sender Email"),
					(self.sender_designation_str,"Sender Designation"),
					(self.sender_department_str,"Sender Department"),
					(self.org_name_str,"Organization Name"),
					(self.addr_txt,"Organization Address"),
					(self.org_website_str,"Organization Website"),
					(self.org_logo_str,"Organization Logo")
				]
		if not self.validate_fields(fields):
			return False
		else:
			dataMap = {
				"sender_name" : self.sender_name_str.get(),
				"sender_email" : self.sender_email_str.get(),
				"sender_designation" : self.sender_designation_str.get(),
				"sender_department" : self.sender_department_str.get(),
				"org_name" : self.org_name_str.get(),
				"org_group" : self.org_group_str.get(),
				"org_addr" : self.addr_txt.get('1.0','end-1c'),
				"org_website" : self.org_website_str.get(),
				"org_logo" : self.org_logo_str.get()
			}
			try:
				self.setup.dumpJSON(dataMap)
				print("Configuration Dumped")
				self.show_alert("showinfo","Configuration Updated")
				self.lift_parent()
				self.root.destroy()
				return
			except Exception:
				self.show_alert("showerror","Unable to Configure, check for file permissions")

if __name__ == "__main__":
	InitSetup()