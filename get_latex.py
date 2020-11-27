import pandas as pd
import docx2txt
import regex
from docx import Document
import pypandoc



def list_to_csv(li,colname,subname):
    """used to convert given list to CSV format

    Args:
        li (list): list of questions with errors
        colname (string): name of college
        subname (string): name of subject
    """
    df = pd.DataFrame(li)
    df.to_csv('errors_'+colname+"_"+subname+'.csv', index=False)



def doc_to_excel(path):
    """used to convert incoming docx file to source latex and then extract required data and put into excel format

    Args:
        path (string): path of file(name of file if in same directory)

    Returns:
        list: list of indices where error has occured
    """


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
            print(temp)
            main.append(temp)
        else:
            errors.append("Question no: "+qno)

    #creates dataframe of the extracted table data        
    df=pd.DataFrame(main,columns=column_names)
    
    #if error occured anywhere then saves the source latex file and a csv containing questions that have error
    if len(errors)>0:
        temps=college_name+"_"+subject_name+".tex"
        #output2 = pypandoc.convert_file(path, 'latex', outputfile=temps)
        list_to_csv(errors,college_name,subject_name)
    
    #converts dataframe to excel format
    #df.to_excel(college_name+"_"+subject_name+".xlsx",index=False)


    return errors 
