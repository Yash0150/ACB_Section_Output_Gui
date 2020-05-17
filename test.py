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
                  
                self.e = Entry(root, width=20, fg='black', 
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
    print(f,h)
    return f>h

def populate(frame,data):
    '''Put in some fake data'''
    for i in range(len(data)):
        a=data[i][0]
        b=data[i][1]
        c=data[i][2]
        l=[b,c,a]
        for j in range(3):
            e = Entry(frame, width=20, fg='black',font=('Arial',16,'bold')) 
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

    data = set()

    for i in range(df.shape[0]):
        x=df['Section'][i].split('-')
        y=x[1][1:]
        y=y+x[1][0]
        if(x[0]==text):
            z=(y,df['StudentID'][i],df['StudentName'][i])
            data.add(z)
    data=list(data)
    data.sort()
    #sorted(data,key=functools.cmp_to_key(mycmp))

    y=("","","")
    z=set()
    i=0
    while(i<len(data)-1):
        #print(data[i][0]!=data[i+1][0])
        if(data[i][0]!=data[i+1][0]):
            data.insert(int(i+1),y)
            i=i+2
            #print(data[i][0],data[i+1])
            z.add(i)
        else:
        	i=i+1

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



