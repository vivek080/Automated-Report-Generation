import random
import time
import csv
import pandas as pd
from docxtpl import DocxTemplate
#import tkinter as tk
from tkinter import * 
from tkinter.ttk import *
from tkinter.filedialog import askopenfilename
# from datetime import datetime as dt
# import datetime
#remove comment from below import lines if you are generating pdf file
# from reportlab.pdfgen import canvas
# from os import chdir, getcwd, listdir, path
# from win32com import client
  
root = Tk() 
root.geometry('500x300')
root.title("Automated Report Generation")

  
# This function will be used to open 

def quit_root():
    root.destroy()

def open_file():
    xf = askopenfilename() #select file from GUI window

    if len(xf) == 0:
        print("File Not Selected! Please rerun the program")
        wait = time.sleep(5)
        exit(0)

    def genrate_file(n):
        doc_file = DocxTemplate("DOH SAMPLE.docx") # word template should be in the same directory
        excel_file = pd.read_excel(xf)
        excel_to_doct = excel_file.to_dict() # dataframe -> dict for the template render
        x = excel_file.to_dict(orient='records')
        ab = x[n]

        #to convert a date column containing timestamp format to date only formate
        # date1 = ab["DISBURSEMENT_DATE"]
        # ab["DISBURSEMENT_DATE"] = date1.date().strftime('%m.%d.%y')

        #print(ab)
        context = ab
        doc_file.render(context)
        a = context
        #print(a)
        doc_file.save("file/%s_report.docx" %a["NAME"]) #to save report as word file
        
        #to save report as pdf
        # tpl.save("file/temp.docx")
        # files = "file/temp.docx"
        # word = client.DispatchEx("Word.Application")
        # new_name = files.replace(".docx", r".pdf")
        # in_file = path.abspath(files)
        # new_file = path.abspath("file/%s_report" %a["NAME"])
        # doc = word.Documents.Open(in_file)
        # print ("Generating pdf ",n)
        # doc.SaveAs(new_file, FileFormat = 17)
        # doc.Close()
        
        

    #--------Main Function---------#   

    file_length = len(pd.read_excel(xf))
    print ("There will be ", file_length, "Reports genrated in File folder")
    wait = time.sleep(2)
    
    for i in range(0,file_length):
        print("Generating file: ",i+1,"Please Wait...")
        genrate_file(i)
        
    print("Report generation complete, check your Reports in File folder!!")
    #root.destroy()
  
Label1 = Label(root)
_img1 = PhotoImage(file="logo.png")
Label1.configure(image=_img1)
Label1.configure(text='''Label''')
Label1.pack()
btn = Button(root, text ='Upload Excel File', command = lambda:open_file()) 
btn.pack(side = TOP, pady = 10)
btn1 = Button(root, text ='Close', command = lambda:quit_root()) 
btn1.pack(side = TOP, pady = 10)
label2 = Label(root, text="Devloped By: VIVEK HIREMATH")
label2.pack(side = BOTTOM, pady = 10)

mainloop() 
