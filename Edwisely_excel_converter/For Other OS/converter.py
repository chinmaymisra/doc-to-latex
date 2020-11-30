import tkinter as tk
from tkinter import *

import pandas as pd
import docx2txt
import regex
from docx import Document
import pypandoc

from tkinter.filedialog import askdirectory

import os



def parse_folder(directory):
    if directory=="ERROR" or directory =="":
        popup = tk.Tk()
        popup.wm_title("ERROR !!")
        errlabel=tk.Label(popup,text="PLEASE SELECT A FOLDER FIRST",fg="RED",font=('helvetica', 20, 'bold'))
        errlabel.pack(side="top", fill="x", pady=10)
        B1 = ttk.Button(popup, text="Okay", command = popup.destroy)
        B1.pack()
        popup.mainloop()
        #errlabel=tk.Label(root,text="PLEASE SELECT A FOLDER FIRST",fg="RED",font=('helvetica', 20, 'bold'))
        #errlabel.grid(row=8)
    directory = directory
    file_li=[]
    for entry in os.scandir(directory):
        if (entry.path.endswith(".docx")):
            file_li.append(entry.path)
    print(file_li)
    return file_li



def list_to_csv(li,colname,subname,path):
    """used to convert given list to CSV format

    Args:
        li (list): list of questions with errors
        colname (string): name of college
        subname (string): name of subject
    """
    df = pd.DataFrame(li)
    df.to_csv(path+'errors_'+colname+"_"+subname+'.csv', index=False)



def doc_to_excel(path):
    """used to convert incoming docx file to source latex and then extract required data and put into excel format

    Args:
        path (string): path of file(name of file if in same directory)

    Returns:
        list: list of indices where error has occured
    """
    path=path.replace("\\","/")
    dirpath=extract(path)
    print("dirpath ="+dirpath)
    
    if path=="ERROR":
        print("SELECT A FILE FIRST")

        
    output = pypandoc.convert_file(path, 'latex')
    l=output.split("\\tabularnewline\r\n")
    size=len(l)
    

    column_names=['Question','Question_image','Question_type','Option1',
              'Option1_image','Option2','Option2_image','Option3','Option3_image',
              'Option4','Option4_image','Option5','Option5_image','Answer','Solution',
              'Solution_image','Hint','Hint_image','Blooms_level','Difficulty_level',
              'Topic_tags','Prerequisite_tags','Source','Concept_Summary_id','Order',
              'math_type']
    
    #college name extraction
    col=regex.findall(r'College name(.*?)College name',l[0])
    col=str(col[0])
    college_name=col.replace("\'","")
    
    #subject name extraction
    sub=regex.findall(r'Subject name(.*?)Subject name',l[0])
    sub=str(sub[0])
    subject_name=sub.replace("\'","")

    #top row
    l[1]=l[1].replace("\\midrule\r\n\\endhead\r\n","")
    
    main=[]
    errors=[]
    for i in range(1,size-1):
        temp=l[i].split("&")
        qno=temp[0]
        temp.pop(0)
        
        if len(temp)==6:
            temp.insert(1,"")
            temp.insert(2,"")
            temp.insert(4,"")
            temp.insert(6,"")
            temp.insert(8,"")
            temp.insert(10,"")
            temp.insert(11,"")
            temp.insert(12,"")
            temp.insert(14,"")
            temp.insert(15,"")
            temp.insert(16,"")
            temp.insert(17,"")
            temp.insert(18,"")
            temp.insert(19,"")
            temp.insert(20,"")
            temp.insert(21,"")
            temp.insert(22,"")
            temp.insert(23,"")
            temp.insert(24,"")
            temp.insert(25,"")
            main.append(temp)
        else:
            errors.append("Question no: "+qno)

    #creates dataframe of the extracted table data        
    df=pd.DataFrame(main,columns=column_names)
    
    #if error occured anywhere then saves the source latex file and a csv containing questions that have error
    if len(errors)>0:
        temps=dirpath+college_name+"_"+subject_name+".tex"
        output2 = pypandoc.convert_file(path, 'latex', outputfile=temps)
        list_to_csv(errors,college_name,subject_name,dirpath)
    
    #converts dataframe to excel format
    df.to_excel(dirpath+college_name+"_"+subject_name+".xlsx",index=False)




    return errors 


def go_convert(li):
    for i in li:
        print("processing :"+i)
        doc_to_excel(i)
    
        







Fname="ERROR"


def choose():
    Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
    filename = askdirectory() # show an "Open" dialog box and return the path to the selected file
    print(filename)
    global Fname
    Fname=filename
    return filename


root= tk.Tk()

root.wm_title("Edwisely Excel Generator                   created by Chinmay")

thelabel=tk.Label(root,text="Edwisely Excel Generator",fg="Blue",font=('helvetica', 12, 'bold'))
thelabel.grid(row=0)

label_2=tk.Label(root,text="Steps to follow:",fg="black")
label_2.grid(row=2,sticky=W)

label_3=tk.Label(root,text="1)Click on \"Select Folder\"",fg="black")
label_3.grid(row=4,sticky=W)

label_4=tk.Label(root,text="2)Find the folder where you have stored the docx files",fg="black")
label_4.grid(row=5,sticky=W)

label_5=tk.Label(root,text="3)Click on \"Convert\"",fg="black")
label_5.grid(row=6,sticky=W)

label_1=tk.Label(root,text="4)Enjoy !",fg="black")
label_1.grid(row=7,sticky=W)




label_1=tk.Label(root,text="Created by Chinmay",fg="red")
label_1.grid(row=1,sticky=E)




canvas1 = tk.Canvas(root, width = 300, height = 300)
canvas1.grid()

    
button1 = tk.Button(root,text='Select Folder',command=(lambda :choose()), bg='green',fg='white')
button1.grid(sticky=S)


button2 = tk.Button(root,text='Convert',command=(lambda:go_convert(parse_folder(Fname))),bg='brown',fg='white')
button2.grid(sticky=S)
#canvas1.create_window(150, 150, window=button1)

def extract(fname):
    l=fname.split("/")
    docname=(l[-1])
    directory="/".join(l[:-1])
    directory+="/"
    print(directory)
    return directory

root.mainloop()
