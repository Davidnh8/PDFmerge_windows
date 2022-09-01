from PyPDF2 import PdfFileMerger, PdfFileReader
import os
from tkinter import *
from tkinter import filedialog, scrolledtext
import datetime

# Window is a child class that inherits from Frame
class Window(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.master = master
        self.init_window()
        self.pdf_name_list=[]
        
        # main text box
        self.textbox = scrolledtext.ScrolledText(self, width=55, height=25)
        self.textbox.pack(anchor='w')
        self.textbox.place(x=5, y=25)
        self.count=1
        
        #top textbox
        self.toptext = Text(self, width=20, height=1, bg='light gray')
        self.toptext.insert(INSERT, "list of pdf")
        self.toptext.pack(anchor='w')
        self.toptext.place(x=5,y=5)
        
        # error textbox
        self.errtext = Text(self, width=55, height=2, bg='light gray')
        self.errtext.pack(anchor='w')
        self.errtext.place(x=5,y=450)
        
        
    def init_window(self):
        self.master.title("PDFMerger")
        self.pack(fill=BOTH, expand=1)
        
        exit_button = Button(self, text="exit", command=root.destroy)
        exit_button.place(relx=0.98, rely=0.9, anchor=E)
        
        add_button = Button(self, text="Add pdf", command=self.add_pdf)
        add_button.place(relx=0.98, rely=0.1, anchor=E, width=85)
        
        remove_button = Button(self, text="reset pdf list", command=self.reset_pdf)
        remove_button.place(relx=0.98, rely=0.2, anchor=E, width=85)
        
        merge_button = Button(self, text="Merge!", command=self.merge_pdf)
        merge_button.place(relx=0.98, rely=0.7, anchor=E, width=85, height=85)
    
    def add_pdf(self):
        self.errtext.delete('1.0', END)
        pdf_file = filedialog.askopenfilenames(parent=self, initialdir=os.getcwd(), filetypes=(("pdf files", "*.pdf"),("all files", "*.*")))
        if (type(pdf_file)==type([])) + (type(pdf_file)==type((1,))):
            
            # error check loop
            for file in pdf_file:
                ext=file.split('.')[-1]
                if (ext=='pdf')+(ext=='PDF')+(ext=='Pdf'):
                    pass
                else:
                    error_message="At least one of the chosen file(s) is not a pdf.\nType detected: "+file.split('.')[-1]
                    self.errtext.insert(INSERT, error_message)
                    raise ValueError(error_message)
            
            # add pdf files
            for file in pdf_file:
                self.pdf_name_list.append(file)
                display = str(self.count)+ ". " + file.split("/")[-1] + "\n"
                self.textbox.insert(INSERT, display)
                self.count+=1
        else:
            raise TypeError("output of filedialog.askopenfilenames is neither list of tuple. Either nothing was chosen, or unknown error has occured")
    def reset_pdf(self):
        self.errtext.delete('1.0', END)
        self.textbox.delete('1.0', END)
        self.pdf_name_list=[]
        self.count=1
    
    def merge_pdf(self):
        self.errtext.delete('1.0', END)
        if len(self.pdf_name_list)==0:
            self.errtext.insert(INSERT, "There is nothing to merge!")
            raise ValueError("There is nothing to merge!")
        elif len(self.pdf_name_list)==1:
            self.errtext.insert(INSERT, "Only single file is present. No need to merge.")
            raise ValueError("Only single file is present. No need to merge.")
        else:
            merger = PdfFileMerger()
            for filename in self.pdf_name_list:
                merger.append(open(filename, 'rb'))
            dt = datetime.datetime.now()
            strdt = str(dt)
            date = strdt.split(" ")[0]
            time = strdt.split(" ")[1]
            time = time.split(":")
            timehr = time[0]
            timemin = time[1]
            timesec = time[2].split(".")[0]
            finaldt = date + " " + timehr + "-" + timemin + "-" + timesec
            with open("mergedPDF_%s.pdf" % finaldt , 'wb') as fout:
                merger.write(fout)
            self.errtext.insert(INSERT, "Merge Complete!!!")
    
root =Tk()
root.geometry("600x500")
app = Window(root)
root.mainloop()