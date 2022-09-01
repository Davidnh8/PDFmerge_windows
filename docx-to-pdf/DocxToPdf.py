from PyPDF2 import PdfFileMerger, PdfFileReader
import os
from tkinter import *
from tkinter import filedialog, scrolledtext
import datetime
import docx
import comtypes.client

# Window is a child class that inherits from Frame
class Window(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.master = master
        self.init_window()
        self.docx_list=[]
        
        # main text box
        self.textbox = scrolledtext.ScrolledText(self, width=55, height=25)
        self.textbox.pack(anchor='w')
        self.textbox.place(x=5, y=25)
        self.count=1
        
        #top textbox
        self.toptext = Text(self, width=20, height=1, bg='light gray')
        self.toptext.insert(INSERT, "list of docx")
        self.toptext.pack(anchor='w')
        self.toptext.place(x=5,y=5)
        
        # error textbox
        self.errtext = Text(self, width=55, height=2, bg='light gray')
        self.errtext.pack(anchor='w')
        self.errtext.place(x=5,y=450)
        
        
    def init_window(self):
        self.master.title("DOCX to PDF")
        self.pack(fill=BOTH, expand=1)
        
        exit_button = Button(self, text="exit", command=root.destroy)
        exit_button.place(relx=0.98, rely=0.9, anchor=E)
        
        add_button = Button(self, text="Add Docx", command=self.add_pdf)
        add_button.place(relx=0.98, rely=0.1, anchor=E, width=85)
        
        remove_button = Button(self, text="reset docx list", command=self.reset_docx)
        remove_button.place(relx=0.98, rely=0.2, anchor=E, width=85)
        
        merge_button = Button(self, text="Convert!", command=self.convert_pdf)
        merge_button.place(relx=0.98, rely=0.7, anchor=E, width=85, height=85)
    
    def add_pdf(self):
        self.errtext.delete('1.0', END)
        pdf_file = filedialog.askopenfilenames(parent=self, initialdir=os.getcwd(), filetypes=(("docx files", "*.docx"),("all files", "*.*")))
        if (type(pdf_file)==type([])) + (type(pdf_file)==type((1,))):
            
            # error check loop
            for file in pdf_file:
                ext=file.split('.')[-1]
                if (ext=='docx')+(ext=='DOCX')+(ext=='Docx'):
                    pass
                else:
                    error_message="At least one of the chosen file(s) is not a docx.\nType detected: "+file.split('.')[-1]
                    self.errtext.insert(INSERT, error_message)
                    raise ValueError(error_message)
            
            # add pdf files
            for file in pdf_file:
                self.docx_list.append(file)
                display = str(self.count)+ ". " + file.split("/")[-1] + "\n"
                self.textbox.insert(INSERT, display)
                self.count+=1
        else:
            raise TypeError("output of filedialog.askopenfilenames is neither list of tuple. Either nothing was chosen, or unknown error has occured")
    def reset_docx(self):
        self.errtext.delete('1.0', END)
        self.textbox.delete('1.0', END)
        self.docx_list=[]
        self.count=1
    
    def convert_pdf(self):
        self.errtext.delete('1.0', END)
        if len(self.docx_list)==0:
            self.errtext.insert(INSERT, "There is nothing to convert!")
            raise ValueError("There is nothing to merge!")
        else:
            word = None
            doc = None
            for i in range(0, len(self.docx_list)):
                word = comtypes.client.CreateObject('Word.Application')
                in_filename = self.docx_list[i].replace('/','\\')
                doc = word.Documents.Open(in_filename)
				
                out_filename = in_filename.split(".")[0] + '.pdf'
                doc.SaveAs(out_filename, FileFormat=17)
                doc.Close()
            word.Quit()
            
            self.errtext.insert(INSERT, "Convert Complete!!!")
    
root =Tk()
root.geometry("600x500")
app = Window(root)
root.mainloop()