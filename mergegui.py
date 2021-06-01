from tkinter import *
from tkinter.ttk import *
from tkinter.scrolledtext import ScrolledText
from tkinter.filedialog import askopenfilename,askdirectory
from tkinter import messagebox

from setup import Setup
from mail_merge import MailMerge
from tkPDFViewer import tkPDFViewer as pdf
from mailgui import MailGUI
from pdfgui import PDF_Preview

import threading
import os
import logging
logger = logging.getLogger(f"MailMerge.{os.path.basename(__file__)}")

class MergeGUI:
    def __init__(self):
        root = Tk()
        root.wm_title("Mail Merge Utility")
        root.minsize(500,200)

        self.template_str = StringVar()
        self.excel_str = StringVar()

        self.root = root

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.resizable(False,False)

        logger.info("Initialising Template Input")
        self.add_template_inp()
        logger.info("Initialising Merge Field display")
        self.add_merge_fields()
        logger.info("Initialising Excel Input")
        self.add_excel_inp()
        logger.info("Initialising Excel Field display")
        self.add_excel_fields()
        logger.info("Initialising Generate Button")
        self.add_generate_btn()
        logger.info("Initialising copyright Label")
        self.add_copyright_lbl()
        

        self.root.mainloop()

    def on_closing(self):
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            self.root.destroy()
            self.root.quit()
            logger.info("Quiting Application")

    def show_alert(self,code,message):
        if code == "showinfo":
            messagebox.showinfo(code,message)
            logger.info(f"{message}")
            return
        messagebox.showerror(code,message)
        logger.error(f"{message}")
    
    def add_template_inp(self):
        self.template_lbl = Label(self.root,text = "Template File")
        self.template_lbl.grid(row = 1, padx = 10 , pady = 10, column = 0, columnspan = 2,sticky = N+S+E+W)
        self.template_dir = Entry(self.root,font = "Consolas 12",textvariable = self.template_str,width = "30",state="readonly")
        self.template_dir.grid(row = 1, column = 3,padx = 10, pady = 10,columnspan = 15,sticky = N+S+E+W)
        self.template_browse = Button(self.root,text = "Browse",command = lambda:self.open_template(self.template_str))
        self.template_browse.grid(row = 1, padx = 10, pady = 10, column = 20, columnspan = 3,sticky = N+S+E+W)

    def add_merge_fields(self):
        self.merge_fields_lbl = Label(self.root,text = "Merge Fields")
        self.merge_fields = ScrolledText(self.root,font = ("Times New Roman",10), state = "disabled",width = "30",height = "1")

    def add_excel_inp(self):
        self.excel_lbl = Label(self.root,text = "Data file")
        self.excel_lbl.grid(row = 3, padx = 10 , pady = 10, column = 0, columnspan = 2,sticky = N+S+E+W)
        self.excel_dir = Entry(self.root,font = "Consolas 12",textvariable = self.excel_str,width = "30",state="readonly")
        self.excel_dir.grid(row = 3, column = 3,padx = 10, pady = 10,columnspan = 15,sticky = N+S+E+W)
        self.excel_browse = Button(self.root,text = "Browse",command = lambda:self.open_excel(self.excel_str))
        self.excel_browse.grid(row = 3, padx = 10, pady = 10, column = 20, columnspan = 3,sticky = N+S+E+W)

    def add_excel_fields(self):
        self.excel_header_lbl = Label(self.root,text = "Data Fields")
        self.excel_headers = ScrolledText(self.root,font = ("Times New Roman",10), state = "disabled",width = "30",height = "1")

    def add_generate_btn(self):
        self.gen = Button(self.root,text = "Start Mail Merge",command = lambda:self.run_script(self.template_str,self.excel_str))
        self.gen.grid(row = 7, padx = 10, pady = 10, column = 2, columnspan = 18,sticky = N+S+E+W)

    def add_copyright_lbl(self):
        copyright = "© 2021 Pratik Goel, Published in India"
        self.cp_lbl = Label(self.root,text = copyright,)
        self.cp_lbl.grid(row = 17, padx = 10 , pady = 10, column = 5, columnspan = 12,sticky = N+S+E+W)

    def add_progress_bar(self):
        self.var = StringVar()
        self.progress_lbl = Label(self.root,textvariable = self.var)
        self.progress_lbl.grid(row = 13, padx = 10 , pady = 10, column = 1, columnspan = 10,sticky = N+S+E+W)

        self.progress = Progressbar(self.root,orient = HORIZONTAL, length = 100, mode = 'determinate')
        self.progress.grid(row = 15, padx = 10, pady = 10, column = 1, columnspan = 60,sticky = N+S+E+W)
        self.progress['value'] = 0

    def run_script(self,template_dir,excel_dir):
        if len(template_dir.get()) < 1:
            self.show_alert("showerror","Enter Template file")
            return
        if len(excel_dir.get()) < 1:
            self.show_alert("showerror","Enter Data file")
            return
        logger.info("Initialising Progressbar")
        self.add_progress_bar()
        
        thread1 = threading.Thread(target = self.start_merge)
        thread1.start()

    def start_merge(self):
        mailMerge = MailMerge(template_file=self.template_dir.get(),excel_file=self.excel_dir.get())
        mailMerge.init_setup()
        mailMerge.read_template()
        mailMerge.read_excel()
        
        data = mailMerge.EXCEL_DATA.read_data()
        total = len(data.index)
        cnt = 70 // total
        idx = 1
        
        for fields in data.to_dict(orient='records'):
            self.progress['value'] += cnt
            self.var.set(f"Merging ({idx} / {total})")
            logger.info(f"Merging ({idx} / {total})")
            idx += 1
            mailMerge.merge_doc(fields)
        
        self.var.set("Converting PDFs")
        logger.info("Converting Merged Docs to PDF")
        
        mailMerge.convert_pdf()
        self.var.set(f"Converted {total} files to PDF")
        logger.info(f"Converted {total} files to PDF")
        
        self.progress['value'] = 100
        self.show_alert("showinfo",f"Processed {total} files")

        logger.info("Removing Progressbar")
        self.progress_lbl.grid_forget()
        self.progress.grid_forget()

        logger.info("Generating PDF Previews")
        PDF_Preview(data)

    def data_str(self,data):
        return ", ".join(sorted(data))

    def open_template(self,template_str):
        filename = askopenfilename(filetypes = [('Microsoft Word','*.docx')])
        self.template_str.set(filename)
        if filename == "":
            self.merge_fields_lbl.grid_remove()
            self.merge_fields.grid_remove()
            return
        doc = MailMerge(template_file=filename).read_template()
        fields = doc.merge_fields()
        fields_str = self.data_str(fields)
        
        logger.info("Displaying merge fields from template file")
        self.merge_fields_lbl.grid(row = 2, padx = 10 , pady = 5, column = 0, columnspan = 2,sticky = N+S+E+W)
        self.merge_fields.configure(state='normal')
        self.merge_fields.delete("1.0",END)
        self.merge_fields.insert(END, fields_str)
        self.merge_fields.configure(state='disabled')
        self.merge_fields.grid(row = 2, column = 3,padx = 10, pady = 5,columnspan = 20,sticky = N+S+E+W)

    def open_excel(self,excel_str):
        filename = askopenfilename(filetypes = [('Microsoft Excel','*.xlsx')])
        self.excel_str.set(filename)
        if filename == "":
            self.excel_header_lbl.grid_remove()
            self.excel_headers.grid_remove()
            return
        excel_obj = MailMerge(excel_file=filename).read_excel()
        excel_data = excel_obj.read_data()
        excel_headers_str = self.data_str(excel_data.columns)
        excel_headers_str += f"\nNumber of Rows : {len(excel_data.index)}"
        
        logger.info("Displaying Excel Headers")
        self.excel_header_lbl.grid(row = 4, padx = 10 , pady = 5, column = 0, columnspan = 2,sticky = N+S+E+W)
        self.excel_headers.configure(state='normal')
        self.excel_headers.delete("1.0",END)
        self.excel_headers.insert(END, excel_headers_str)
        self.excel_headers.configure(state='disabled')
        self.excel_headers.grid(row = 4, column = 3,padx = 10, pady = 5,columnspan = 20,sticky = N+S+E+W)
            
