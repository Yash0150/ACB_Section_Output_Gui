from tkinter import *
import tkinter as tk
from tkinter import ttk
from PIL import Image,ImageTk
import os
import pandas as pd
import functools
import xlrd
import os
import sys
import pandas as pd
import numpy as np
import xlsxwriter 


class Table: 
      
    def __init__(self,root,data): 
          
        # code for creating table 
       
        for i in range(len(data)):
            
            for j in range(3): 
                  
                self.e = Entry(root, width=20, fg='blue', 
                               font=('Arial',10,'bold')) 
                  
                self.e.grid(row=i, column=j)
                if(j==2): 
                    self.e.insert(END, data[i][j]) 

def mycmp(s1, s2):
    (a,b,c)=s1
    (x,y,z)=s2
    f=c[1:]
    h=z[1:]
    f=str(f)+str(c[0])
    h=str(h)+str(z[0])
    #print(f,h)
    return f>h

def populate(frame,data):
    '''Put in some fake data'''
    for i in range(len(data)):
        a=data[i][0]
        b=data[i][1]
        c=data[i][2]
        l=[b,c,a]
        for j in range(3):
            e = Entry(frame, width=20, fg='blue',font=('Arial',16,'bold')) 
            e.grid(row=i, column=j)
            if(j==2): 
                x=len(a)
                y=""
                if(x!=0):
                    y=a[x-1]+a[0:x-1]
                if(x>8):
                    y='       SECTION'
                e.insert(END, y)
            else: 
                e.insert(END,l[j]) 


def onFrameConfigure(canvas):
    '''Reset the scroll region to encompass the inner frame'''
    canvas.configure(scrollregion=canvas.bbox("all"))

def show():
    text=clicked.get()
    sec=clicked2.get()
    data = set()
    for i in range(df.shape[0]):
        x=df['Section'][i].split('-')
        y=x[1][1:]
        y=y+x[1][0]
        if(x[0]==text and sec==x[1]):
            z=(y,df['StudentID'][i],df['StudentName'][i])
            #print(sec==x[1])
            data.add(z)
    data=list(data)
    data.sort()
    #sorted(data,key=functools.cmp_to_key(mycmp))

#     y=("","","")
#     z=set()
#     for i in range(0,len(data)-1):
#         print(data[i][0]!=data[i+1][0])
#         if(data[i][0]!=data[i+1][0]):
#             data.insert(int(i+1)+1,y)
#             i=i+1
#             print(data[i][0],data[i+1])
#             z.add(i)
    c=0;
    # data2=set()
    # for i in data:
    #     (x,q,w)=i;
    #     # a=x[0]
    #     # b=x[1:]+a
    #     # x=b
    #     data2.add((q,w,x))
    # data=list(data2)
    # for i in z:
    #     print(i)
    #     if(c==0):
    #     	c=0
    #     data.insert(c+int(i)+1,y)
    #     if(c==-1):
    #     	c=0
    #     c=c+1;
    y=('       SECTION','      ID','       NAME')
    data.insert(0,y)
    y=("","","")
    data.insert(1,y)
    root = tk.Tk()
    root.geometry("700x400") 
    canvas = tk.Canvas(root, borderwidth=0, background="#ffffff")
    frame = tk.Frame(canvas, background="#ffffff")
    vsb = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=vsb.set)

    vsb.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    canvas.create_window((4,4), window=frame, anchor="nw")

    frame.bind("<Configure>", lambda event, canvas=canvas: onFrameConfigure(canvas))

    populate(frame,data)

#     root.mainloop() 

def gen():
    text=clicked.get()
    sec=clicked2.get()
    workbook = xlsxwriter.Workbook(text+'-'+sec+'.xlsx')
    workbook.add_worksheet(text+sec)
    cell_format = workbook.add_format({'bold': True, 'align': 'center'})
    cell_format3 =workbook.add_format ({'align': 'center'})
    worksheet=workbook.get_worksheet_by_name(text+sec)
    worksheet.set_column(0,200, 50)
    worksheet.write(0,0,'NAME',cell_format)
    worksheet.write(0,1,'ID',cell_format)
    worksheet.write(0,2,'SECTION',cell_format)
    data = set()
    for i in range(df.shape[0]):
        x=df['Section'][i].split('-')
        y=x[1][1:]
        y=y+x[1][0]
        if(x[0]==text and sec==x[1]):
            z=(y,df['StudentID'][i],df['StudentName'][i])
            #print(sec==x[1])
            data.add(z)
    data=list(data)
    data.sort()
    y=('       SECTION','      ID','       NAME')
    data.insert(0,y)
    y=("","","")
    data.insert(1,y)
    for i in range(len(data)):
        a=data[i][0]
        b=data[i][1]
        c=data[i][2]
        l=[b,c,a]
        for j in range(3):
            if(j==2): 
                x=len(a)
                y=""
                if(x!=0):
                    y=a[x-1]+a[0:x-1]
                if(x>8):
                    y='       SECTION'
                worksheet.write(i,j,y,cell_format3)
            else: 
                worksheet.write(i,j,l[j],cell_format3)
    workbook.close() 

def stcount():
    Dict= {}
    Dict_CN={}
    #print(df2['Subject'])
    for i in range(df2.shape[0]):
        a=df2['Subject'][i]+df2['Catalog'][i]
        a=a.replace(" ","")
        Dict_CN[a]=df2['Course Title'][i]
    for i in range(df.shape[0]):
        x=df['Section'][i]
        y=df['StudentName'][i]
        z=df['StudentID'][i]
        #w=df['CourseName'][i]
        Dict[x]=0
        #Dict_CN[x.split('-')[0]]=w
    for i in range(df.shape[0]):
        x=df['Section'][i]
        Dict[x]+=1
#     print(Dict_CN)
    workbook = xlsxwriter.Workbook('stdcount'+'.xlsx')
    workbook.add_worksheet('count')
    cell_format = workbook.add_format({'bold': True, 'align': 'center'})
    cell_format3 =workbook.add_format ({'align': 'center'})
    worksheet=workbook.get_worksheet_by_name('count')
    worksheet.set_column(0,200, 50)
    worksheet.write(0,0,'COURSE',cell_format)
    worksheet.write(0,1,'SECTION',cell_format)
    worksheet.write(0,2,'COURSE NAME',cell_format)
    worksheet.write(0,3,'COUNT',cell_format)
#     data = set()
    cnt=2
    for k in Dict:
       # print(k.split('-'))
        worksheet.write(cnt,0,k.split('-')[0],cell_format3)
        worksheet.write(cnt,1,k.split('-')[1],cell_format3)
        worksheet.write(cnt,2,Dict_CN[k.split('-')[0]],cell_format3)
        worksheet.write(cnt,3,Dict[k],cell_format3)
        cnt+=1
    workbook.close() 
    
    
p_dir = os.getcwd()
df=pd.read_excel('Course_List.xls')
df2=pd.read_excel('Timetable.xlsx',header=1)
course=set()

for i in range(df.shape[0]):
    x=df['Section'][i].split('-')
    course.add(x[0])

course=list(course)

Sections=[]

for i in range(10):
    Sections.insert(0,'P'+str(i))
    Sections.insert(0,'L'+str(i))
    Sections.insert(0,'T'+str(i))


root = Tk()
root.title('ACB-INFO')
root.geometry("900x600") 



my_img = ImageTk.PhotoImage(Image.open("logo.png"))
mylable = Label(image=my_img)
mylable.pack()

mylable=Label(root,text="PLEASE SELECT THE COURSE",pady=20).pack()

clicked = StringVar()
clicked.set(course[0])

course.sort()

drop = ttk.Combobox(root, textvariable=clicked, values=course)
drop.pack()

clicked2 = StringVar()
clicked2.set(Sections[0])

Sections.sort()

drop2 = ttk.Combobox(root, textvariable=clicked2, values=Sections)
drop2.pack()

myButton = Button(root, text="Show Result", command=show,pady=20)
myButton.pack()

myButton2 = Button(root, text="Generate Excel", command=gen,pady=20)
myButton2.pack()

myButton3 = Button(root, text="Summary", command=stcount,pady=20)
myButton3.pack()

root.mainloop()



