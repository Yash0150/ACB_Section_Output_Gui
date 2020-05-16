from tkinter import *
import tkinter as tk
from tkinter import ttk
from PIL import Image,ImageTk
import os
import pandas as pd
import functools

class Table: 
      
    def __init__(self,root,data): 
          
        # code for creating table 
        for i in range(len(data)): 
            for j in range(3): 
                  
                self.e = Entry(root, width=20, fg='blue', 
                               font=('Arial',10,'bold')) 
                  
                self.e.grid(row=i, column=j) 
                self.e.insert(END, data[i][j]) 

def mycmp(s1, s2):
    (a,b,c)=s1
    (x,y,z)=s2
    f=c[1:]
    h=z[1:]
    f=str(f)+str(c[0])
    h=str(h)+str(z[0])
    #print(f,h)
    return f>=h

def populate(frame,data):
    '''Put in some fake data'''
    for i in range(len(data)): 
        for j in range(3):
            e = Entry(frame, width=20, fg='blue',font=('Arial',16,'bold')) 
            e.grid(row=i, column=j)
            e.insert(END, data[i][j]) 


def onFrameConfigure(canvas):
    '''Reset the scroll region to encompass the inner frame'''
    canvas.configure(scrollregion=canvas.bbox("all"))

def show():
    text=clicked.get()

    data = set()

    for i in range(df.shape[0]):
        x=df['Section'][i].split('-')
        if(x[0]==text):
            y=(df['StudentID'][i],df['StudentName'][i],x[1])
            data.add(y)
    data=list(data)

    sorted(data,key=functools.cmp_to_key(mycmp))

    y=("","","")
    z=set()
    for i in range(0,len(data)-1):
        #print(data[i][2][1]!=data[i+1][2][1],data[i][2][1],data[i+1][2][1])
        if(data[i][2][1:]!=data[i+1][2][1:] or data[i][2][0]!=data[i+1][2][0]):
            z.add(i+1)
    c=0;
    for i in z:
        data.insert(c+int(i),y)
        c=c+1;
    y=('      ID','       NAME','       SECTION')
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

    root.mainloop() 


p_dir = os.getcwd()
df=pd.read_excel('Course_List.xls')
course=set()

for i in range(df.shape[0]):
    x=df['Section'][i].split('-')
    course.add(x[0])

course=list(course)




root = Tk()
root.title('ACB-INFO')
root.geometry("700x400") 



my_img = ImageTk.PhotoImage(Image.open("logo.png"))
mylable = Label(image=my_img)
mylable.pack()

mylable=Label(root,text="PLEASE SELECT THE COURSE",pady=20).pack()

clicked = StringVar()
clicked.set(course[0])

course.sort()

drop = ttk.Combobox(root, textvariable=clicked, values=course)
drop.pack()

myButton = Button(root, text="Show Result", command=show,pady=20)
myButton.pack()


root.mainloop()



