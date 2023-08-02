# -*- coding: utf-8 -*-
"""
Created on Sat Aug  7 13:32:34 2021

@author: IFEANYI
"""
import os
import shapefile
import tkinter.filedialog
import tkinter as tk
from tkinter import messagebox, ttk
from tkinter import *
import math
import pandas as pd
from pandas import DataFrame
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from pandastable import Table
import ezdxf
from PIL import Image, ImageTk
import sys
import pyttsx3

#---------------------------Define the GUI--------------------------------------
master = Tk()
master.geometry('700x400')
master['bg'] = 'powder blue'
master.title('Survey Companion')

icon = tk.PhotoImage(file = 'S_Companion Icon.png')
master.iconphoto(False,icon)


def User_guide():
    try: del(sys.modules["UserGuide"])
    except:""
    import UserGuide
   
def about():
    try: del(sys.modules["About"])
    except:""
    import About

frame0 = Frame(master,relief='groove',bg='powder blue') 
frame0.pack(fill='both', expand='yes')

class Resize(Frame):
    def __init__(self, master, *pargs):
        Frame.__init__(self, master, *pargs)

        self.image = Image.open('S_Companion landing page.jpg')
        self.img_copy= self.image.copy()


        self.background_image = ImageTk.PhotoImage(self.image)

        self.background = Label(self, image=self.background_image)
        self.background.pack(fill='both', expand='yes')
        self.background.bind('<Configure>', self._resize_image)

    def _resize_image(self,event):

        new_width = event.width
        new_height = event.height

        self.image = self.img_copy.resize((new_width, new_height))

        self.background_image = ImageTk.PhotoImage(self.image)
        self.background.configure(image =  self.background_image)
home = Resize(frame0)
home.pack(fill='both', expand='yes')

opendatalabel = Label(frame0,font=('Times', 16, 'bold'),fg = 'red')
opendatalabel['text'] = 'No file has been opened yet'
opendatalabel.place(relx=0.0,rely=0.9)

# Traverse Home Page


def onFrameConfigure(canvas): canvas.configure(scrollregion=canvas.bbox("all")) 
        
frametraverse0 = Frame(master,bd=5,relief='groove',bg='powder blue') 

canvastraverse = Canvas(frametraverse0, borderwidth=0, background='powder blue')
frametraverse00 = Frame(canvastraverse,bg='powder blue') 

vsb = Scrollbar(frametraverse0, orient="vertical", command=canvastraverse.yview)
vsb.pack(side="right", fill="y")
hsb = Scrollbar(frametraverse0, orient="horizontal", command=canvastraverse.xview)
hsb.pack(side="bottom", fill="x")
canvastraverse.pack(side="left", expand=True, fill="both")

frametraverse00.bind("<Configure>", lambda event, canvas=canvastraverse: onFrameConfigure(canvastraverse))
canvastraverse.create_window((0,0), window=frametraverse00, anchor="nw")
canvastraverse.configure(yscrollcommand=vsb.set)
canvastraverse.configure(xscrollcommand=hsb.set)


frametraverse1 = LabelFrame(frametraverse00, text='Notes for your data file!!!',bg = 'powder blue',bd=5,padx=10,pady=10 , fg="red2",font=('times new roman', 12, 'bold','italic'))   
labeltraverse1 = Label(frametraverse1, text='''Instructions for Users!!!

(i) This program adjusts a closed traverse network(both loop 
and connecting traverses)
(ii) Open your data file from the file menu in the Menu bar 
(iii) Ensure your data file is saved properly using the .xlsx 
file extension eg. Datafile.xlsx
(iv) For your excel data files, fill in the Station_At, Face,
Horizontal Angles, Distance and Station_To (in this order) in 
different columns (one for each).
(v) Your Angular observations have to be separated with a space 
not the conventional degrees, minutes and seconds symbols.
(vi) Also provide the Starting and Closing Controls and their
coordinates
(vii) To show a plot of the traverse network or the computation 
Sheet, kindly check the corresponding boxes.
(viii) Click the 'Compute' button to proceed with the adjustment. 
''')
labeltraverse1.config(font=('helvetica', 9,'bold','italic'),bg = 'CadetBlue1', justify = 'left')
labeltraverse1.grid(row=0,column=0,sticky='nw')

frametraverse4 = LabelFrame(frametraverse00, text='Open File!!!',bg = 'powder blue',padx=10,pady=10 , font=('times new roman', 12, 'bold','italic'),relief = 'flat')

labeltraverse6 = Label(frametraverse4,font=('Times', 12, 'bold'),bg = 'powder blue',fg = 'red')
labeltraverse6['text'] = 'No file Selected'
labeltraverse6.grid()      
frametraverse5 = LabelFrame(frametraverse00,bg = 'powder blue',bd=5,padx=108,pady=10)
checkvartraverse = IntVar()
plottraverse = Checkbutton(frametraverse5, text='Plot the Traverse Network',variable=checkvartraverse,bg = 'powder blue', cursor='hand2') 
plottraverse.grid(row=0)
checkvartraverse1 = IntVar()
sheet = Checkbutton(frametraverse5, text='Show Computation Sheet',variable=checkvartraverse1,bg = 'powder blue', cursor='hand2') 
sheet.grid(row=1)

labeltraversemethod = Label(frametraverse5,text='Choose Computation Method:',font=('Times', 12, 'bold'),bg = 'powder blue')
traversemethod = StringVar()
traverseoption = ttk.Combobox(frametraverse5,width=27,textvariable=traversemethod, values=('Choose Computation Method:',"Bowditch Rule","Transit Rule"), cursor='hand2',state='readonly')
traverseoption.current(0)

    
traverseoption.grid(row=2, column = 0, pady=10)

# =============================================================================
# Level Home Page
framelevel0 = Frame(master,bd=5,relief='groove',bg='powder blue') 

canvaslevel = Canvas(framelevel0, borderwidth=0, background='powder blue')
framelevel00 = Frame(canvaslevel,bg='powder blue') 

vsb = Scrollbar(framelevel0, orient="vertical", command=canvaslevel.yview)
vsb.pack(side="right", fill="y")
hsb = Scrollbar(framelevel0, orient="horizontal", command=canvaslevel.xview)
hsb.pack(side="bottom", fill="x")
canvaslevel.pack(side="left", expand=True, fill="both")

framelevel00.bind("<Configure>", lambda event, canvas=canvaslevel: onFrameConfigure(canvaslevel))
canvaslevel.create_window((0,0), window=framelevel00, anchor="nw")
canvaslevel.configure(yscrollcommand=vsb.set)
canvaslevel.configure(xscrollcommand=hsb.set)

framelevel1 = LabelFrame(framelevel00, text='Notes for your data file!!!',bg = 'powder blue',bd=5,padx=10,pady=10 , fg="red2",font=('times new roman', 12, 'bold','italic'))   
labellevel1 = Label(framelevel1, text='''Instructions for Users!!!

(i) This program adjusts a closed level network(both loop 
and connecting networks)
(ii) Open your data file from the file menu in the Menu bar
(iii) Ensure your data file is saved properly using the .xlsx 
file extension eg. Datafile.xlsx
(iii) For your excel data files, fill in the Staff_Stations, 
Chainages, Backsights, Intermediate sight and Foresights (in
this order) in different columns (one for each).
(iv) Make sure to include the closing benchmark in the 
Staff_Stations column upon level closure.
(v) Also provide the Starting and Closing Benchmark and their 
Reduced heights
(v) To show the reduction table, kindly check the corresponding box.
(vi) Click the 'Reduce' button to proceed with the Reduction and 
corrections. ''')
labellevel1.config(font=('helvetica', 9,'bold','italic'),bg = 'CadetBlue1', justify = 'left')
labellevel1.grid(row=0,column=0,sticky='nw')

framelevel4 = LabelFrame(framelevel00, text='Open File!!!',bg = 'powder blue',padx=10,pady=2 , font=('times new roman', 12, 'bold','italic'),relief = 'flat')

labellevel6 = Label(framelevel4,font=('Times', 12, 'bold'),bg = 'powder blue',fg = 'red')
labellevel6['text'] = 'No file Selected'
labellevel6.grid(columnspan=3)      
framelevel5 = LabelFrame(framelevel00,bg = 'powder blue',bd=5,padx=125,)

checkvarlevel1 = IntVar()
sheetlevel = Checkbutton(framelevel5, text='Show Computation Sheet',variable=checkvarlevel1,bg = 'powder blue', cursor='hand2') 
sheetlevel.grid()

coordframelevel = LabelFrame(framelevel00,text='Corrected Reduced Levels',bg='CadetBlue1',bd=2, padx=10,pady=10,relief=RIDGE, fg="green4",font=('Times', 14, 'bold', 'italic'))

# =============================================================================
# Hydro Home Page
framehydro0 = Frame(master,bd=5,relief='groove',bg='powder blue')
 
canvashydro = Canvas(framehydro0, borderwidth=0, background='powder blue')
framehydro00 = Frame(canvashydro,bg='powder blue') 

vsb = Scrollbar(framehydro0, orient="vertical", command=canvashydro.yview)
vsb.pack(side="right", fill="y")
hsb = Scrollbar(framehydro0, orient="horizontal", command=canvashydro.xview)
hsb.pack(side="bottom", fill="x")
canvashydro.pack(side="left", expand=True, fill="both")

framehydro00.bind("<Configure>", lambda event, canvas=canvashydro: onFrameConfigure(canvashydro))
canvashydro.create_window((0,0), window=framehydro00, anchor="nw")
canvashydro.configure(yscrollcommand=vsb.set)
canvashydro.configure(xscrollcommand=hsb.set)

framehydro1 = LabelFrame(framehydro00, text='Notes for your data file!!!',bg = 'powder blue',bd=5,padx=10,pady=10 , fg="red2",font=('times new roman', 12, 'bold','italic'))   
labelhydro1 = Label(framehydro1, text='''Instructions for Users!!!

(i) This program reduces sounding(Bathymetric) data and
filters it based on set conditions
(ii) Open your data file from the file menu in the Menu bar 
(iii) Ensure your data file is saved properly using the .csv 
file extension eg. Datafile.csv
(iv) For your data files, fill in the Time, Eastings,
Northings and Depths (in this order) 
(v) Your Time values(hh:mm:ss) have to be separated with colon(:) 
eg. 10:05:47
(vi) You will also be required to import your tide data file, also 
saved in .csv file extension eg. Tidefile.csv
(vii) For your tide data files, fill in the Time(hh:mm:ss) and 
Tide levels (in this order).
(viii) To show a plot of the spot depths or the contours,
kindly check the corresponding boxes.
(ix) Click the 'Reduce' button to proceed with the reduction 
and filtering. 
''')
labelhydro1.config(font=('helvetica', 9,'bold','italic'),bg = 'CadetBlue1', justify = 'left')
labelhydro1.grid(row=0,column=0,sticky='nw')

framehydro2 = LabelFrame(framehydro00, text='Data Filter', bg='CadetBlue1',bd=5,font=('times new roman', 12, 'bold','italic'),pady=3)
framehydro2.grid(row=0,column=1,sticky='nw')

labelhydro2 = Label(framehydro2, text='\nEnter tolerable depth range below:',bg = 'CadetBlue1',font=('times new roman', 12, 'bold'),justify = 'left')
labelhydro2.grid(row=0,column=0,columnspan=3,sticky = 'w')

labelhydro3 = Label(framehydro2, text=' Greater than » ',bg = 'CadetBlue1', font=('times new roman', 12, 'bold'), pady=5)
labelhydro4 = Label(framehydro2, text='And less than » ',bg = 'CadetBlue1', font=('times new roman', 12, 'bold'), pady=5)

def on_click_hydro1(event):
    greater_than.config(foreground='black')
    if greater_than.get() == "0.З":event.widget.delete(0, tk.END)   
    else:greater_than.config(foreground='black')
myvar_hydro1 = tk.StringVar()
myvar_hydro1.set("0.З")
greater_than = Entry(framehydro2,font=('Times', 12, 'bold'), width=10, relief=RIDGE, bd=5, textvariable=myvar_hydro1)
greater_than.config(foreground='gray80')
greater_than.bind("<Button-1>", on_click_hydro1)
greater_than.bind("<FocusIn>", on_click_hydro1)

def on_click_hydro2(event):
    less_than.config(foreground='black')
    if less_than.get() == "З0":event.widget.delete(0, tk.END)   
    else:less_than.config(foreground='black')
myvar_hydro2 = tk.StringVar()
myvar_hydro2.set("З0")
less_than = Entry(framehydro2,font=('Times', 12, 'bold'), width=10, relief=RIDGE, bd=5, textvariable=myvar_hydro2)
less_than.config(foreground='gray80')
less_than.bind("<Button-1>", on_click_hydro2)
less_than.bind("<FocusIn>", on_click_hydro2)

labelhydro3unit = Label(framehydro2, text='m',bg = 'CadetBlue1', font=('times new roman', 12, 'bold', 'italic'), pady=5, fg='gray50')
labelhydro4unit = Label(framehydro2, text='m',bg = 'CadetBlue1', font=('times new roman', 12, 'bold', 'italic'), pady=5, fg='gray50')

labelhydro9 = Label(framehydro2, text='\n\nFilter data by distance interval below:',bg = 'CadetBlue1',font=('times new roman', 12, 'bold'),justify = 'left')
labelhydro9.grid(row=3,column=0,columnspan=3,sticky = 'w')
labelhydro10 = Label(framehydro2, text='Distance Interval » ',bg = 'CadetBlue1', font=('times new roman', 12, 'bold'), pady=5)

def on_click_hydro3(event):
    distance_interval.config(foreground='black')
    if distance_interval.get() == "З":event.widget.delete(0, tk.END)   
    else:distance_interval.config(foreground='black')
myvar_hydro3 = tk.StringVar()
myvar_hydro3.set("З")
distance_interval = Entry(framehydro2,font=('Times', 12, 'bold'), width=10, relief=RIDGE, bd=5, textvariable=myvar_hydro3)
distance_interval.config(foreground='gray80')
distance_interval.bind("<Button-1>", on_click_hydro3)
distance_interval.bind("<FocusIn>", on_click_hydro3)

labelhydro10unit = Label(framehydro2, text='m',bg = 'CadetBlue1', font=('times new roman', 12, 'bold', 'italic'), pady=5, fg='gray50')

labelhydro3.grid(row=1,column=0,sticky = 'e')
labelhydro4.grid(row=2,column=0,sticky = 'e')
greater_than.grid(row=1,column=1)
less_than.grid(row=2,column=1)
labelhydro3unit.grid(row=1,column=2,sticky = 'w')
labelhydro4unit.grid(row=2,column=2,sticky = 'w')
labelhydro10.grid(row=4,column=0,sticky = 'e')
distance_interval.grid(row=4,column=1)
labelhydro10unit.grid(row=4,column=2,sticky = 'w')

framehydro4 = LabelFrame(framehydro00, text='Open File!!!',bg = 'powder blue',pady=10 , font=('times new roman', 12, 'bold','italic'),relief = 'flat')
framehydro6 = LabelFrame(framehydro00, text='Open Tide Observation File!!!',bg = 'powder blue',pady=10 , font=('times new roman', 12, 'bold','italic'),relief = 'flat')

labelhydro7 = Label(framehydro2, text='\n\n\nClick "OPEN" to open tide data file',bg = 'CadetBlue1',font=('times new roman', 12, 'bold'),justify = 'left')

labelhydro6 = Label(framehydro4,font=('Times', 12, 'bold'),bg = 'powder blue',fg = 'red')
labelhydro6['text'] = 'No file Selected'
labelhydro6.grid(columnspan=3)  

labelhydro8 = Label(framehydro6,font=('Times', 12, 'bold'),bg = 'powder blue',fg = 'red')
labelhydro8['text'] = 'No file Selected'
labelhydro8.grid(columnspan=3)
  
framehydro5 = LabelFrame(framehydro00,bg = 'powder blue',bd=5,pady=20,padx=275)
checkvarhydro = IntVar()
plothydro = Checkbutton(framehydro5, text='Plot the Spot Depth',variable=checkvarhydro,bg = 'powder blue', cursor='hand2') 
plothydro.grid()

framehydro7 = LabelFrame(framehydro00, text="**Filtered Coordinates**",bd=5, bg="CadetBlue1", fg="green4",font=('Times', 14, 'bold', 'italic'))

# =============================================================================
# Area Home Page
framearea0 = Frame(master,bd=5,relief='groove',bg='powder blue') 

canvasarea = Canvas(framearea0, borderwidth=0, background='powder blue')
framearea00 = Frame(canvasarea,bg='powder blue') 

vsb = Scrollbar(framearea0, orient="vertical", command=canvasarea.yview)
vsb.pack(side="right", fill="y")
hsb = Scrollbar(framearea0, orient="horizontal", command=canvasarea.xview)
hsb.pack(side="bottom", fill="x")
canvasarea.pack(side="left", expand=True, fill="both")

framearea00.bind("<Configure>", lambda event, canvas=canvasarea: onFrameConfigure(canvasarea))
canvasarea.create_window((0,0), window=framearea00, anchor="nw")
canvasarea.configure(yscrollcommand=vsb.set)
canvasarea.configure(xscrollcommand=hsb.set)

framearea1 = LabelFrame(framearea00, text='Notes for your data file/input!!!',bg = 'powder blue',bd=5,padx=10,pady=13 , fg="red2",font=('times new roman', 12, 'bold','italic'))   
framearea1.grid(row=0,column=0,sticky=NW)
labelarea1 = Label(framearea1, text='''**Instructions for Users!!!

(i) This program computes the area of an enclosed region
using the coordinates of the vertices of the region.
(ii) This program either accepts coordinate inputs directly
or reads coordinates from files in .xlsx, .csv and .txt file
extensions.
(iii) The coordinates should be filled in consecutively either
in a clockwise or anti-clockwise manner without repeating the
first point.
(iv) For excel files in .xlsx or .csv extensions, enter the
coordinates in two differnt columns, eastings(x) and northings(y)
(one for each).
(v) For direct input and text files in .txt extension, separate
the coordinates (easting and northing) with a comma eg. 1234,9876.
Each vertex coordinate in a new line.
(vi) To show a plot of the enclosed region, kindly check the box.
(vii) Click the 'Compute' button to calculate the area.
 ''')
labelarea1.config(font=('helvetica', 9,'bold','italic'),padx=25,bg = 'CadetBlue1', justify = 'left')
labelarea1.grid(row=0,column=0)


framearea2 = LabelFrame(framearea00, text='Data Input', bg='CadetBlue1',bd=5,font=('times new roman', 12, 'bold','italic'))

labelarea4 = Label(framearea2, text='Enter your coordinates below: ',bg = 'CadetBlue1')
labelarea4.config(font=('times new roman', 12, 'bold'))
labelarea4.grid(row=0,column=0,sticky='w')

scrollbarinput = Scrollbar(framearea2)
entry = Text(framearea2,height=18, width=25)#, textvariable=myvar_area)


entry.grid(row=1,column=0,sticky='nw')
scrollbarinput.grid(row=1,column=1,sticky='ns')

scrollbarinput.config(command=entry.yview)

framearea4 = LabelFrame(framearea00, text='Open File!!!',bg = 'powder blue',pady=10 , font=('times new roman', 12, 'bold','italic'),relief='flat')
labelarea6 = Label(framearea4,font=('Times', 12, 'bold'),bg = 'powder blue',fg = 'red')
labelarea6['text'] = 'No file Selected'
labelarea6.grid(columnspan=3)

framearea5 = LabelFrame(framearea00,bg = 'powder blue',bd=5,padx=267,pady=10)
framearea5.grid(row=2,column=0,columnspan=2)

frameareaoutput = LabelFrame(framearea00,text='Computed Results',bg='CadetBlue1',bd=2, padx=10,pady=10,relief=RIDGE, fg="green4",font=('Times', 14, 'bold', 'italic'))

scrollbaroutput = Scrollbar(frameareaoutput)
labelareaoutput = Text(frameareaoutput, width=40, yscrollcommand=scrollbaroutput.set)

scrollbaroutput.grid(row=0,column=1,sticky='ns')
checkvararea = IntVar()
plot = Checkbutton(framearea5, text='Plot the Enclosed Region',variable=checkvararea,bg = 'powder blue', cursor='hand2') 
plot1 = Checkbutton(framearea5, text='Plot the Enclosed Region',variable=checkvararea,bg = 'powder blue', cursor='hand2')  

scrollbaroutput.config(command=labelareaoutput.yview)

# =============================================================================
# Resection Home Page
frameresection0 = Frame(master,bd=5,relief='groove',bg='powder blue') 

canvasresection = Canvas(frameresection0, borderwidth=0, background='powder blue')
frameresection00 = Frame(canvasresection,bg='powder blue') 

vsb = Scrollbar(frameresection0, orient="vertical", command=canvasresection.yview)
vsb.pack(side="right", fill="y")
hsb = Scrollbar(frameresection0, orient="horizontal", command=canvasresection.xview)
hsb.pack(side="bottom", fill="x")
canvasresection.pack(side="left", expand=True, fill="both")

frameresection00.bind("<Configure>", lambda event, canvas=canvasresection: onFrameConfigure(canvasresection))
canvasresection.create_window((0,0), window=frameresection00, anchor="nw")
canvasresection.configure(yscrollcommand=vsb.set)
canvasresection.configure(xscrollcommand=hsb.set)

frameresection1 = LabelFrame(frameresection00, text='Notes for your data input!!!',bg = 'powder blue',bd=5,padx=10,pady=10 , fg="red2",font=('times new roman', 12, 'bold','italic'))   
labelresection1 = Label(frameresection1, text='''**Instructions for Users!!!

(i) This program computes the coordinate of an unknown station     
using the three point resection method of position fixing
(ii) The program accepts direct inputs of the control 
coordinates needed for the computation.
(iii) The coordinates should be entered in easting and 
northing pairs, and separated by a comma, eg.
123456.789,987654.321 
(iv) Also, the interior angle subtended at the unknown point 
by points A and B and points B and C are required.
(v) Angular inputs should be numeric and in the form dd mm ss,  
with the degree, minutes and seconds values separated by 
space eg. 25 06 14
(vi) Note that all angular observations are assumed clockwise,i.e. 
Angle AÔB is assumed the clockwise angle between the directions 
OA and OB
(vii) Click the 'Compute' button to calculate the coordinate
of the resected point.
''')
labelresection1.config(font=('helvetica', 9,'bold','italic'),padx=25,bg = 'CadetBlue1', justify = 'left')
labelresection1.grid(row=0,column=0)

frameresection2 = LabelFrame(frameresection00, text='Data Input', bg='CadetBlue1',bd=5,font=('times new roman', 12, 'bold','italic'),pady=8)
labelresection4 = Label(frameresection2, text='''                        
Enter Control Station Coordinates(x,y) below: ''',bg = 'CadetBlue1')
labelresection4.config(font=('times new roman', 12, 'bold'))
labelresection4.grid(row=0,column=0,columnspan=3,sticky='w')

labelresection5 = Label(frameresection2, text='         Control_A » ',bg = 'CadetBlue1', font=('times new roman', 12, 'bold'), pady=5)
labelresection6 = Label(frameresection2, text='         Control_B » ',bg = 'CadetBlue1', font=('times new roman', 12, 'bold'), pady=5)
labelresection7 = Label(frameresection2, text='         Control_C » ',bg = 'CadetBlue1', font=('times new roman', 12, 'bold'), pady=5)


def on_click_resection1(event):
    Control_A.config(foreground='black')
    if Control_A.get() == "12З456.789,987654.З21":event.widget.delete(0, tk.END)   
    else:Control_A.config(foreground='black')
myvar_resection1 = tk.StringVar()
myvar_resection1.set("12З456.789,987654.З21")
Control_A = Entry(frameresection2,font=('Times', 12, 'bold'), width=20, relief=RIDGE,bd=5, textvariable=myvar_resection1)
Control_A.config(foreground='gray80')
Control_A.bind("<Button-1>", on_click_resection1)
Control_A.bind("<FocusIn>", on_click_resection1)

def on_click_resection2(event):
    Control_B.config(foreground='black')
    if Control_B.get() == "12З456.789,987654.З21":event.widget.delete(0, tk.END)   
    else:Control_B.config(foreground='black')
myvar_resection2 = tk.StringVar()
myvar_resection2.set("12З456.789,987654.З21")
Control_B = Entry(frameresection2,font=('Times', 12, 'bold'), width=20, relief=RIDGE,bd=5, textvariable=myvar_resection2)
Control_B.config(foreground='gray80')
Control_B.bind("<Button-1>", on_click_resection2)
Control_B.bind("<FocusIn>", on_click_resection2)

def on_click_resection3(event):
    Control_C.config(foreground='black')
    if Control_C.get() == "12З456.789,987654.З21":event.widget.delete(0, tk.END)   
    else:Control_C.config(foreground='black')
myvar_resection3 = tk.StringVar()
myvar_resection3.set("12З456.789,987654.З21")
Control_C = Entry(frameresection2,font=('Times', 12, 'bold'), width=20, relief=RIDGE,bd=5, textvariable=myvar_resection3)
Control_C.config(foreground='gray80')
Control_C.bind("<Button-1>", on_click_resection3)
Control_C.bind("<FocusIn>", on_click_resection3)



labelresection5unit = Label(frameresection2, text='mE,mN',bg = 'CadetBlue1', font=('times new roman', 12, 'bold','italic'), pady=5, fg='gray50')
labelresection6unit = Label(frameresection2, text='mE,mN',bg = 'CadetBlue1', font=('times new roman', 12, 'bold','italic'), pady=5, fg='gray50')
labelresection7unit = Label(frameresection2, text='mE,mN',bg = 'CadetBlue1', font=('times new roman', 12, 'bold','italic'), pady=5, fg='gray50')


labelresection5.grid(row=1,column=0,sticky='e')
labelresection6.grid(row=2,column=0,sticky='e')
labelresection7.grid(row=3,column=0,sticky='e')
Control_A.grid(row=1, column=1)
Control_B.grid(row=2, column=1)
Control_C.grid(row=3, column=1)
labelresection5unit.grid(row=1,column=2,sticky='w')
labelresection6unit.grid(row=2,column=2,sticky='w')
labelresection7unit.grid(row=3,column=2,sticky='w')

frameresection5 = LabelFrame(frameresection00,bg = 'powder blue',bd=5,padx=332,pady=20)
frameresection5.grid(row=2,column=0,columnspan=2)
labelresection8 = Label(frameresection2, text='\n\n\n\nEnter Angular Observations(dd mm ss) below:',bg = 'CadetBlue1')
labelresection8.config(font=('times new roman', 12, 'bold'))
labelresection8.grid(row=5,column=0,columnspan=3,sticky='w')

labelresection9 = Label(frameresection2, text='   Angle AÔB » ',bg = 'CadetBlue1', font=('times new roman', 12, 'bold'), pady=5)
labelresection10 = Label(frameresection2, text='   Angle BÔC » ',bg = 'CadetBlue1', font=('times new roman', 12, 'bold'), pady=5)

def on_click_resection4(event):
    Angle_AÔB.config(foreground='black')
    if Angle_AÔB.get() == "dd mm ss":event.widget.delete(0, tk.END)   
    else:Angle_AÔB.config(foreground='black')
myvar_resection4 = tk.StringVar()
myvar_resection4.set("dd mm ss")
Angle_AÔB = Entry(frameresection2,font=('Times', 12, 'bold'), width=20, relief=RIDGE,bd=5, textvariable=myvar_resection4)
Angle_AÔB.config(foreground='gray80')
Angle_AÔB.bind("<Button-1>", on_click_resection4)
Angle_AÔB.bind("<FocusIn>", on_click_resection4)

def on_click_resection5(event):
    Angle_BÔC.config(foreground='black')
    if Angle_BÔC.get() == "dd mm ss":event.widget.delete(0, tk.END)   
    else:Angle_BÔC.config(foreground='black')
myvar_resection5 = tk.StringVar()
myvar_resection5.set("dd mm ss")
Angle_BÔC = Entry(frameresection2,font=('Times', 12, 'bold'), width=20, relief=RIDGE,bd=5, textvariable=myvar_resection5)
Angle_BÔC.config(foreground='gray80')
Angle_BÔC.bind("<Button-1>", on_click_resection5)
Angle_BÔC.bind("<FocusIn>", on_click_resection5)

labelresection9unit = Label(frameresection2, text=' °  \'  \"',bg = 'CadetBlue1', font=('times new roman', 12, 'bold','italic'), pady=5, fg='gray50')
labelresection10unit = Label(frameresection2, text=' °  \'  \"',bg = 'CadetBlue1', font=('times new roman', 12, 'bold','italic'), pady=5, fg='gray50')

labelresection9.grid(row=6,column=0,sticky='e')
labelresection10.grid(row=7,column=0,sticky='e')

Angle_AÔB.grid(row=6, column=1)
Angle_BÔC.grid(row=7, column=1)

labelresection9unit.grid(row=6,column=2,sticky='w')
labelresection10unit.grid(row=7,column=2,sticky='w')

frameresectionoutput = LabelFrame(frameresection00,text='Computed Results',bg='CadetBlue1',bd=2, padx=10,pady=10,relief=RIDGE, fg="green4",font=('Times', 14, 'bold', 'italic'))
labelresectionoutput = Text(frameresectionoutput, width=50,height = 20)

checkvarresection = IntVar()
plotresection = Checkbutton(frameresection5, text='Plot the Resected point',variable=checkvarresection,bg = 'powder blue', cursor='hand2') 
plotresection.grid()

# =============================================================================
# Intersection Home Page
frameintersection0 = Frame(master,bd=5,relief='groove',bg='powder blue') 

canvasintersection = Canvas(frameintersection0, borderwidth=0, background='powder blue')
frameintersection00 = Frame(canvasintersection,bg='powder blue') 

vsb = Scrollbar(frameintersection0, orient="vertical", command=canvasintersection.yview)
vsb.pack(side="right", fill="y")
hsb = Scrollbar(frameintersection0, orient="horizontal", command=canvasintersection.xview)
hsb.pack(side="bottom", fill="x")
canvasintersection.pack(side="left", expand=True, fill="both")

frameintersection00.bind("<Configure>", lambda event, canvas=canvasintersection: onFrameConfigure(canvasintersection))
canvasintersection.create_window((0,0), window=frameintersection00, anchor="nw")
canvasintersection.configure(yscrollcommand=vsb.set)
canvasintersection.configure(xscrollcommand=hsb.set)


frameintersection1 = LabelFrame(frameintersection00, text='Notes for your data inputs!!!',bg = 'powder blue',bd=5,padx=10,pady=10 , fg="red2",font=('times new roman', 12, 'bold','italic'))   
frameintersection1.grid(row=0,column=0,sticky=NW)
labelintersection1 = Label(frameintersection1, text='''**Instructions for Users!!!

(i) This program computes the coordinate of an unknown station     
using the intersection method of position fixing
(ii) The program accepts direct inputs of the control 
coordinates needed for the computation.
(iii) The coordinates should be entered in easting and 
northing pairs, and separated by a comma, eg.
123456.789,987654.321 
(iv) Depending on the type of intersection, you will be required
to fill in the appropriate data values
(v) Angular inputs should be numeric and in the form dd mm ss,  
with the degree, minutes and seconds values separated by 
space eg. 25 06 14. Likewise, distance inputs should be numeric.
(vi) Note that all angular observations are assumed interior 
angles subtended at a point by the triangle formed by the control
points and the unknown point.
(vii) Click the 'Compute' button to calculate the coordinate
of the intersected point.
''')
labelintersection1.config(font=('helvetica', 9,'bold','italic'),padx=25,bg = 'CadetBlue1', justify = 'left')
labelintersection1.grid(row=0,column=0)

frameintersection2 = LabelFrame(frameintersection00, text='Data Input', bg='CadetBlue1',bd=5,font=('times new roman', 12, 'bold','italic'),pady=5)
labelintersection4 = Label(frameintersection2, text='''Enter Control Station Coordinates(x,y) below:   ''',bg = 'CadetBlue1')
labelintersection4.config(font=('times new roman', 12, 'bold'))
labelintersection4.grid(row=0,column=0,columnspan=3,sticky='w')

labelintersection5 = Label(frameintersection2, text='       Control_A » ',bg = 'CadetBlue1', font=('times new roman', 12, 'bold'), pady=5)
labelintersection6 = Label(frameintersection2, text='       Control_B » ',bg = 'CadetBlue1', font=('times new roman', 12, 'bold'), pady=5)

def on_click_intersection1(event):
    Control_A_intersection.config(foreground='black')
    if Control_A_intersection.get() == "12З456.789,987654.З21":event.widget.delete(0, tk.END)   
    else:Control_A_intersection.config(foreground='black')
myvar_intersection1 = tk.StringVar()
myvar_intersection1.set("12З456.789,987654.З21")
Control_A_intersection = Entry(frameintersection2,font=('Times', 12, 'bold'), width=20, relief=RIDGE,bd=5, textvariable=myvar_intersection1)
Control_A_intersection.config(foreground='gray80')
Control_A_intersection.bind("<Button-1>", on_click_intersection1)
Control_A_intersection.bind("<FocusIn>", on_click_intersection1)

def on_click_intersection2(event):
    Control_B_intersection.config(foreground='black')
    if Control_B_intersection.get() == "12З456.789,987654.З21":event.widget.delete(0, tk.END)   
    else:Control_B_intersection.config(foreground='black')
myvar_intersection2 = tk.StringVar()
myvar_intersection2.set("12З456.789,987654.З21")
Control_B_intersection = Entry(frameintersection2,font=('Times', 12, 'bold'), width=20, relief=RIDGE,bd=5, textvariable=myvar_intersection2)
Control_B_intersection.config(foreground='gray80')
Control_B_intersection.bind("<Button-1>", on_click_intersection2)
Control_B_intersection.bind("<FocusIn>", on_click_intersection2)

labelintersection5unit = Label(frameintersection2, text='mE,mN',bg = 'CadetBlue1', font=('times new roman', 12, 'bold','italic'), pady=5, fg='gray50')
labelintersection6unit = Label(frameintersection2, text='mE,mN',bg = 'CadetBlue1', font=('times new roman', 12, 'bold','italic'), pady=5, fg='gray50')

labelintersection5.grid(row=1,column=0,sticky='e')
labelintersection6.grid(row=2,column=0,sticky='e')
Control_A_intersection.grid(row=1, column=1)
Control_B_intersection.grid(row=2, column=1)
labelintersection5unit.grid(row=1,column=2,sticky='w')
labelintersection6unit.grid(row=2,column=2,sticky='w')

frameintersection5 = LabelFrame(frameintersection00,bg = 'powder blue',bd=5,padx=330,pady=20)
frameintersection51 = LabelFrame(frameintersection00,bg = 'powder blue',bd=5,padx=330,pady=20)
frameintersection52 = LabelFrame(frameintersection00,bg = 'powder blue',bd=5,padx=330,pady=20)
frameintersection53 = LabelFrame(frameintersection00,bg = 'powder blue',bd=5,padx=330,pady=20)
frameintersection54 = LabelFrame(frameintersection00,bg = 'powder blue',bd=5,padx=330,pady=20)
frameintersection55 = LabelFrame(frameintersection00,bg = 'powder blue',bd=5,padx=330,pady=20)

labelintersection8 = Label(frameintersection2, text='''\n\n\n\n\n\n\nEnter Azimuth of each line(dd mm ss) below:     ''',bg = 'CadetBlue1',font=('times new roman', 12, 'bold'),justify = 'left')
labelintersection81 = Label(frameintersection2, text='''\n\n\n\n\n\nEnter Azimuth of line AO(dd mm ss)                \nand distance BO(m) below:                               ''',bg = 'CadetBlue1',font=('times new roman', 12, 'bold'),justify = 'left')
labelintersection82 = Label(frameintersection2, text='''\n\nEnter Distances from each control(m) below:    ''',bg = 'CadetBlue1',font=('times new roman', 12, 'bold'),justify = 'left')
labelintersection83 = Label(frameintersection2, text='''\nEnter Interior Angles(dd mm ss) subtended    \nat the controls by the unknown point below:    ''',bg = 'CadetBlue1',font=('times new roman', 12, 'bold'),justify = 'left')
labelintersection84 = Label(frameintersection2, text='''\nEnter Interior Angles(dd mm ss) subtended at\nA and the unknown point by B below:             ''',bg = 'CadetBlue1',font=('times new roman', 12, 'bold'),justify = 'left')
labelintersection85 = Label(frameintersection2, text='''Enter Interior Angle(dd mm ss) subtended at\nthe unknown point and the distance(m)\nfrom A to the unknown point below:''', bg = 'CadetBlue1',font=('times new roman', 12, 'bold'),justify = 'left')

labelintersection9 = Label(frameintersection2, text='   Azimuth AO » ',bg = 'CadetBlue1', font=('times new roman', 12, 'bold'), pady=5)
labelintersection10 = Label(frameintersection2, text='   Azimuth BO » ',bg = 'CadetBlue1', font=('times new roman', 12, 'bold'), pady=5)

labelintersection91 = Label(frameintersection2, text='    Azimuth AO » ',bg = 'CadetBlue1', font=('times new roman', 12, 'bold'), pady=5)
labelintersection101 = Label(frameintersection2, text='   Distance BO » ',bg = 'CadetBlue1', font=('times new roman', 12, 'bold'), pady=5)

labelintersection92 = Label(frameintersection2, text='    Distance AO » ',bg = 'CadetBlue1', font=('times new roman', 12, 'bold'), pady=5)
labelintersection102 = Label(frameintersection2, text='   Distance BO » ',bg = 'CadetBlue1', font=('times new roman', 12, 'bold'), pady=5)

labelintersection93 = Label(frameintersection2, text='   Angle BÂO » ',bg = 'CadetBlue1', font=('times new roman', 12, 'bold'), pady=5)
labelintersection103 = Label(frameintersection2, text='   Angle AḂO » ',bg = 'CadetBlue1', font=('times new roman', 12, 'bold'), pady=5)

labelintersection94 = Label(frameintersection2, text='   Angle BÂO » ',bg = 'CadetBlue1', font=('times new roman', 12, 'bold'), pady=5)
labelintersection104 = Label(frameintersection2, text='   Angle AÔB » ',bg = 'CadetBlue1', font=('times new roman', 12, 'bold'), pady=5)

labelintersection95 = Label(frameintersection2, text='   Angle AÔB » ',bg = 'CadetBlue1', font=('times new roman', 12, 'bold'), pady=5)
labelintersection105 = Label(frameintersection2, text='   Distance AO » ',bg = 'CadetBlue1', font=('times new roman', 12, 'bold'), pady=5)

def on_click_intersection3(event):
    Azimuth_AO_zz_intersection.config(foreground='black')
    if Azimuth_AO_zz_intersection.get() == "dd mm ss":event.widget.delete(0, tk.END)   
    else:Azimuth_AO_zz_intersection.config(foreground='black')
myvar_intersection3 = tk.StringVar()
myvar_intersection3.set("dd mm ss")
Azimuth_AO_zz_intersection = Entry(frameintersection2,font=('Times', 12, 'bold'), width=20, relief=RIDGE,bd=5, textvariable=myvar_intersection3)
Azimuth_AO_zz_intersection.config(foreground='gray80')
Azimuth_AO_zz_intersection.bind("<Button-1>", on_click_intersection3)
Azimuth_AO_zz_intersection.bind("<FocusIn>", on_click_intersection3)

def on_click_intersection4(event):
    Azimuth_BO_zz_intersection.config(foreground='black')
    if Azimuth_BO_zz_intersection.get() == "dd mm ss":event.widget.delete(0, tk.END)   
    else:Azimuth_BO_zz_intersection.config(foreground='black')
myvar_intersection4 = tk.StringVar()
myvar_intersection4.set("dd mm ss")
Azimuth_BO_zz_intersection = Entry(frameintersection2,font=('Times', 12, 'bold'), width=20, relief=RIDGE,bd=5, textvariable=myvar_intersection4)
Azimuth_BO_zz_intersection.config(foreground='gray80')
Azimuth_BO_zz_intersection.bind("<Button-1>", on_click_intersection4)
Azimuth_BO_zz_intersection.bind("<FocusIn>", on_click_intersection4)

def on_click_intersection5(event):
    Azimuth_AO_ad_intersection.config(foreground='black')
    if Azimuth_AO_ad_intersection.get() == "dd mm ss":event.widget.delete(0, tk.END)   
    else:Azimuth_AO_ad_intersection.config(foreground='black')
myvar_intersection5 = tk.StringVar()
myvar_intersection5.set("dd mm ss")
Azimuth_AO_ad_intersection = Entry(frameintersection2,font=('Times', 12, 'bold'), width=20, relief=RIDGE,bd=5, textvariable=myvar_intersection5)
Azimuth_AO_ad_intersection.config(foreground='gray80')
Azimuth_AO_ad_intersection.bind("<Button-1>", on_click_intersection5)
Azimuth_AO_ad_intersection.bind("<FocusIn>", on_click_intersection5)

def on_click_intersection6(event):
    Distance_BO_ad_intersection.config(foreground='black')
    if Distance_BO_ad_intersection.get() == "12З456.789":event.widget.delete(0, tk.END)   
    else:Distance_BO_ad_intersection.config(foreground='black')
myvar_intersection6 = tk.StringVar()
myvar_intersection6.set("12З456.789")
Distance_BO_ad_intersection = Entry(frameintersection2,font=('Times', 12, 'bold'), width=20, relief=RIDGE,bd=5, textvariable=myvar_intersection6)
Distance_BO_ad_intersection.config(foreground='gray80')
Distance_BO_ad_intersection.bind("<Button-1>", on_click_intersection6)
Distance_BO_ad_intersection.bind("<FocusIn>", on_click_intersection6)

def on_click_intersection7(event):
    Distance_AO_dd_intersection.config(foreground='black')
    if Distance_AO_dd_intersection.get() == "12З456.789":event.widget.delete(0, tk.END)   
    else:Distance_AO_dd_intersection.config(foreground='black')
myvar_intersection7 = tk.StringVar()
myvar_intersection7.set("12З456.789")
Distance_AO_dd_intersection = Entry(frameintersection2,font=('Times', 12, 'bold'), width=20, relief=RIDGE,bd=5, textvariable=myvar_intersection7)
Distance_AO_dd_intersection.config(foreground='gray80')
Distance_AO_dd_intersection.bind("<Button-1>", on_click_intersection7)
Distance_AO_dd_intersection.bind("<FocusIn>", on_click_intersection7)

def on_click_intersection8(event):
    Distance_BO_dd_intersection.config(foreground='black')
    if Distance_BO_dd_intersection.get() == "12З456.789":event.widget.delete(0, tk.END)   
    else:Distance_BO_dd_intersection.config(foreground='black')
myvar_intersection8 = tk.StringVar()
myvar_intersection8.set("12З456.789")
Distance_BO_dd_intersection = Entry(frameintersection2,font=('Times', 12, 'bold'), width=20, relief=RIDGE,bd=5, textvariable=myvar_intersection8)
Distance_BO_dd_intersection.config(foreground='gray80')
Distance_BO_dd_intersection.bind("<Button-1>", on_click_intersection8)
Distance_BO_dd_intersection.bind("<FocusIn>", on_click_intersection8)

def on_click_intersection9(event):
    Angle_BAO_aa_intersection.config(foreground='black')
    if Angle_BAO_aa_intersection.get() == "dd mm ss":event.widget.delete(0, tk.END)   
    else:Angle_BAO_aa_intersection.config(foreground='black')
myvar_intersection9 = tk.StringVar()
myvar_intersection9.set("dd mm ss")
Angle_BAO_aa_intersection = Entry(frameintersection2,font=('Times', 12, 'bold'), width=20, relief=RIDGE,bd=5, textvariable=myvar_intersection9)
Angle_BAO_aa_intersection.config(foreground='gray80')
Angle_BAO_aa_intersection.bind("<Button-1>", on_click_intersection9)
Angle_BAO_aa_intersection.bind("<FocusIn>", on_click_intersection9)

def on_click_intersection10(event):
    Angle_ABO_aa_intersection.config(foreground='black')
    if Angle_ABO_aa_intersection.get() == "dd mm ss":event.widget.delete(0, tk.END)   
    else:Angle_ABO_aa_intersection.config(foreground='black')
myvar_intersection10 = tk.StringVar()
myvar_intersection10.set("dd mm ss")
Angle_ABO_aa_intersection = Entry(frameintersection2,font=('Times', 12, 'bold'), width=20, relief=RIDGE,bd=5, textvariable=myvar_intersection10)
Angle_ABO_aa_intersection.config(foreground='gray80')
Angle_ABO_aa_intersection.bind("<Button-1>", on_click_intersection10)
Angle_ABO_aa_intersection.bind("<FocusIn>", on_click_intersection10)

def on_click_intersection11(event):
    Angle_BAO_aa_sidesection.config(foreground='black')
    if Angle_BAO_aa_sidesection.get() == "dd mm ss":event.widget.delete(0, tk.END)   
    else:Angle_BAO_aa_sidesection.config(foreground='black')
myvar_intersection11 = tk.StringVar()
myvar_intersection11.set("dd mm ss")
Angle_BAO_aa_sidesection = Entry(frameintersection2,font=('Times', 12, 'bold'), width=20, relief=RIDGE,bd=5, textvariable=myvar_intersection11)
Angle_BAO_aa_sidesection.config(foreground='gray80')
Angle_BAO_aa_sidesection.bind("<Button-1>", on_click_intersection11)
Angle_BAO_aa_sidesection.bind("<FocusIn>", on_click_intersection11)

def on_click_intersection12(event):
    Angle_AOB_aa_sidesection.config(foreground='black')
    if Angle_AOB_aa_sidesection.get() == "dd mm ss":event.widget.delete(0, tk.END)   
    else:Angle_AOB_aa_sidesection.config(foreground='black')
myvar_intersection12 = tk.StringVar()
myvar_intersection12.set("dd mm ss")
Angle_AOB_aa_sidesection = Entry(frameintersection2,font=('Times', 12, 'bold'), width=20, relief=RIDGE,bd=5, textvariable=myvar_intersection12)
Angle_AOB_aa_sidesection.config(foreground='gray80')
Angle_AOB_aa_sidesection.bind("<Button-1>", on_click_intersection12)
Angle_AOB_aa_sidesection.bind("<FocusIn>", on_click_intersection12)

def on_click_intersection13(event):
    Angle_AOB_ad_sidesection.config(foreground='black')
    if Angle_AOB_ad_sidesection.get() == "dd mm ss":event.widget.delete(0, tk.END)   
    else:Angle_AOB_ad_sidesection.config(foreground='black')
myvar_intersection13 = tk.StringVar()
myvar_intersection13.set("dd mm ss")
Angle_AOB_ad_sidesection = Entry(frameintersection2,font=('Times', 12, 'bold'), width=20, relief=RIDGE,bd=5, textvariable=myvar_intersection13)
Angle_AOB_ad_sidesection.config(foreground='gray80')
Angle_AOB_ad_sidesection.bind("<Button-1>", on_click_intersection13)
Angle_AOB_ad_sidesection.bind("<FocusIn>", on_click_intersection13)

def on_click_intersection14(event):
    Distance_AO_ad_sidesection.config(foreground='black')
    if Distance_AO_ad_sidesection.get() == "12З456.789":event.widget.delete(0, tk.END)   
    else:Distance_AO_ad_sidesection.config(foreground='black')
myvar_intersection14 = tk.StringVar()
myvar_intersection14.set("12З456.789")
Distance_AO_ad_sidesection = Entry(frameintersection2,font=('Times', 12, 'bold'), width=20, relief=RIDGE,bd=5, textvariable=myvar_intersection14)
Distance_AO_ad_sidesection.config(foreground='gray80')
Distance_AO_ad_sidesection.bind("<Button-1>", on_click_intersection14)
Distance_AO_ad_sidesection.bind("<FocusIn>", on_click_intersection14)

labelintersection9unit = Label(frameintersection2, text=' °  \'  \"',bg = 'CadetBlue1', font=('times new roman', 12, 'bold', 'italic'), pady=5, fg='gray50')
labelintersection10unit = Label(frameintersection2, text=' °  \'  \"',bg = 'CadetBlue1', font=('times new roman', 12, 'bold', 'italic'), pady=5, fg='gray50')

labelintersection91unit = Label(frameintersection2, text=' °  \'  \"',bg = 'CadetBlue1', font=('times new roman', 12, 'bold', 'italic'), pady=5, fg='gray50')
labelintersection101unit = Label(frameintersection2, text='m',bg = 'CadetBlue1', font=('times new roman', 12, 'bold', 'italic'), pady=5, fg='gray50')

labelintersection92unit = Label(frameintersection2, text='m',bg = 'CadetBlue1', font=('times new roman', 12, 'bold', 'italic'), pady=5, fg='gray50')
labelintersection102unit = Label(frameintersection2, text='m',bg = 'CadetBlue1', font=('times new roman', 12, 'bold', 'italic'), pady=5, fg='gray50')

labelintersection93unit = Label(frameintersection2, text=' °  \'  \"',bg = 'CadetBlue1', font=('times new roman', 12, 'bold', 'italic'), pady=5, fg='gray50')
labelintersection103unit = Label(frameintersection2, text=' °  \'  \"',bg = 'CadetBlue1', font=('times new roman', 12, 'bold', 'italic'), pady=5, fg='gray50')

labelintersection94unit = Label(frameintersection2, text=' °  \'  \"',bg = 'CadetBlue1', font=('times new roman', 12, 'bold', 'italic'), pady=5, fg='gray50')
labelintersection104unit = Label(frameintersection2, text=' °  \'  \"',bg = 'CadetBlue1', font=('times new roman', 12, 'bold', 'italic'), pady=5, fg='gray50')

labelintersection95unit = Label(frameintersection2, text=' °  \'  \"',bg = 'CadetBlue1', font=('times new roman', 12, 'bold', 'italic'), pady=5, fg='gray50')
labelintersection105unit = Label(frameintersection2, text='m',bg = 'CadetBlue1', font=('times new roman', 12, 'bold', 'italic'), pady=5, fg='gray50')

labelintersection_order = Label(frameintersection2, text='\nWhat is the clockwise order of the triangle formed\nby the controls and the unknown point?',bg = 'CadetBlue1', font=('times new roman', 11, 'bold'), pady=5, justify='left')
ordervar = StringVar()
ordervar.set('Order')
option = OptionMenu(frameintersection2,ordervar,"AOB","BOA")

frameintersectionoutput = LabelFrame(frameintersection00,text='Computed Results',bg='CadetBlue1',bd=2, padx=10,pady=10,relief=RIDGE, fg="green4",font=('Times', 14, 'bold', 'italic'))
labelintersectionoutput = Text(frameintersectionoutput, width=50,height = 22)

checkvarintersection = IntVar()
plotintersection = Checkbutton(frameintersection5, text='Plot the Intersected point',variable=checkvarintersection,bg = 'powder blue', cursor='hand2') 
plotintersection.grid()
plotintersection = Checkbutton(frameintersection51, text='Plot the Intersected point',variable=checkvarintersection,bg = 'powder blue', cursor='hand2') 
plotintersection.grid()
plotintersection = Checkbutton(frameintersection52, text='Plot the Intersected point',variable=checkvarintersection,bg = 'powder blue', cursor='hand2') 
plotintersection.grid()
plotintersection = Checkbutton(frameintersection53, text='Plot the Intersected point',variable=checkvarintersection,bg = 'powder blue', cursor='hand2') 
plotintersection.grid()
plotintersection = Checkbutton(frameintersection54, text='Plot the Intersected point',variable=checkvarintersection,bg = 'powder blue', cursor='hand2') 
plotintersection.grid()
plotintersection = Checkbutton(frameintersection55, text='Plot the Intersected point',variable=checkvarintersection,bg = 'powder blue', cursor='hand2') 
plotintersection.grid()

# =============================================================================
# General Functions

#--------------------Function to Open File dialog box to select data file from your system--------------------
def find():
    master.filename = tk.filedialog.askopenfilename()
    
    labeltraverse6['text'] = master.filename
    labeltraverse6.config(font=('Times', 12, 'bold'),bg = 'powder blue',fg = 'green4')
    frametraverse4.grid(row=1,column=0,columnspan=3,sticky='nw')
    if master.filename == '': 
        labeltraverse6['text'] = 'No file Selected'
        labeltraverse6['fg'] = 'red'
     
    labellevel6['text'] = master.filename
    labellevel6.config(font=('Times', 12, 'bold'),bg = 'powder blue',fg = 'green4')
    framelevel4.grid(row=1,column=0,columnspan=3,sticky='nw')
    if master.filename == '': 
        labellevel6['text'] = 'No file Selected'
        labellevel6['fg'] = 'red'
    
    labelhydro6['text'] = master.filename
    labelhydro6.config(font=('Times', 12, 'bold'),bg = 'powder blue',fg = 'green4')
    framehydro4.grid(row=1,column=0,columnspan=3,sticky='nw')
    if master.filename == '': 
        labelhydro6['text'] = 'No file Selected'
        labelhydro6['fg'] = 'red'
        
    labelarea6['text'] = master.filename
    labelarea6.config(font=('Times', 12, 'bold'),bg = 'powder blue',fg = 'green4')
    framearea4.grid(row=1,column=0,columnspan=3, sticky='nw')
    if master.filename == '': 
        labelarea6['text'] = 'No file Selected'
        labelarea6['fg'] = 'red'
    
    opendatalabel['text'] = 'You opened\n' + master.filename
    opendatalabel.config(font=('Times', 16, 'bold'),fg = 'green4',justify='left')
    if master.filename == '': 
        opendatalabel['text'] = 'You failed to open any data file'
        opendatalabel['fg'] = 'red'
        
    
    return master.filename

def find2():
    master.filename = tk.filedialog.askopenfilename()
    labelhydro8['text'] = master.filename
    labelhydro8.config(font=('Times', 12, 'bold'),bg = 'powder blue',fg = 'green4')
    framehydro6.grid(row=2,column=0,columnspan=3,sticky='nw')    
    if master.filename == '': 
        labelhydro8['text'] = 'No file Selected'
        labelhydro8['fg'] = 'red'
    return master.filename



def Exit():
    exit = messagebox.askyesno('Close all Windows', 'Are you sure of this action?')
    if exit>0:
        try: 
            master.destroy()
            try: 
                window.destroy()
            except:''
        except:''

#----------------Function to compute change in easting(given the distance and bearing)----------------------------        
def easting_compt(distance, bearing):
    departure = distance * math.sin(math.radians(bearing))
    return departure    

#----------------Function to compute change in northing(given the distance and bearing)----------------------------        
def northing_compt(distance, northing):
    latitude = distance * math.cos(math.radians(northing))
    return latitude

#---------------------Function to sum the items in a list(given a list with same type items)--------------------------------
def sum_list(list1):
    sum = 0
    for i in list1:
        sum += i
    return sum

def distance(change_easting,change_northing):
    dist = math.sqrt(change_easting**2+change_northing**2)
    return dist

def distance2(x1,y1,x2,y2):
    dist = math.sqrt((float(x2)-float(x1))**2+(float(y2)-float(y1))**2)
    return dist

def distance3(point1,point2):
    x1,y1 = point1.split(",")
    x2,y2 = point2.split(",")
    dist = math.sqrt((float(x2)-float(x1))**2+(float(y2)-float(y1))**2)
    return dist

def createplot(plot_title):
    window = Toplevel()
    window.title(plot_title)
    window.geometry("600x600")
    window['bg'] = 'powder blue'
    window.iconphoto(False,icon)
    canvas = FigureCanvasTkAgg(fig1, master = window)
    toolbar = NavigationToolbar2Tk(canvas, window)
    toolbar.update()
    toolbar.pack()
    canvas.draw()
    canvas.get_tk_widget().pack()
    

        
#---------------------Function to compute bearing(given change in easting and change in northings)--------------------------------    
def first_bearing(change_easting, change_northing):
    if change_easting > 0 and change_northing > 0:
        B_B = math.atan(change_easting/change_northing)
        B_B = math.degrees(B_B)
        return B_B
        
    if change_easting < 0 and change_northing > 0:
        B_B = math.atan(abs(change_easting)/change_northing)
        B_B = math.degrees(B_B)
        B_B = 360 - B_B
        return B_B
            
    if change_easting < 0 and change_northing < 0:
        B_B = math.atan(abs(change_easting)/abs(change_northing))
        B_B = math.degrees(B_B)
        B_B = 180 + B_B
        return B_B
        
    if change_easting > 0 and change_northing < 0:
        B_B = math.atan(change_easting/abs(change_northing))
        B_B = math.degrees(B_B)
        B_B = 180 - B_B
        return B_B
    
#---------------------Function to convert angle from deg,min,sec to decimal degrees(given the angle as a string in the format 'deg min sec')--------------------------------        
def dms_to_dgs(string):
    angle = str(string)
    degrees, minutes, seconds = angle.split()
    dd = float(degrees) + float(minutes)/60 + float(seconds)/(60*60)
    return dd

def hms_to_dgs(string):
    time = str(string)
    hours, minutes, seconds = time.split(':')
    hh = float(hours) + float(minutes)/60 + float(seconds)/(60*60)
    return hh

def dgs_to_radians(angle):
    return math.radians(angle)

def radians_to_dgs(angle):
    return math.degrees(angle)

#---------------------Function to convert angle from decimal degrees to deg,min,sec(given the angle)--------------------------------     
def dgs_to_dms(angle):
    mnt,sec = divmod(abs(angle)*3600,60)
    deg,mnt = divmod(mnt,60)
    if angle<0:
        return '-'+str(int(deg))+'° '+str(int(mnt))+'\' '+str(float('{:.2f}'.format(sec)))+'\"'
    else:
        return str(int(deg))+'° '+str(int(mnt))+'\' '+str(float('{:.2f}'.format(sec)))+'\"'

#---------------------Function to round up a float to 3 decimal places------------------------------
def rd3(value):
    return float('{:.3f}'.format(value))

#---------------------Function to round up a float to 7 decimal places------------------------------
def rd6(value):
    return float('{:.6f}'.format(value))

def center(widget):
    widget.tag_configure('center', justify='center')
    widget.tag_add('center', 1.0,'end')
    widget.configure(state="disabled")
    
def exportexcel():
    export_file_path = tk.filedialog.asksaveasfilename(defaultextension='.xlsx')
    df4.to_excel (export_file_path, index = False, header=True)
    
def exportcsv():         
    export_file_path = tk.filedialog.asksaveasfilename(defaultextension='.csv')
    df4.to_csv (export_file_path, index = False, header=True)

def exportshp(): 
    try:
        files_to_unhide = ['shapefiletrialpoint.shp','shapefiletrialpoint.shx','shapefiletrialpoint.dbf','shapefiletrialline.shp','shapefiletrialline.shx','shapefiletrialline.dbf']
        for fn in files_to_unhide:
            p = os.popen('attrib -h ' + fn)
            p.read()
            p.close() 
    except: ''
    shp_files = ['shapefiletrialpoint.shp','shapefiletrialline.shp']
    shx_files = ['shapefiletrialpoint.shx','shapefiletrialline.shx']
    dbf_files = ['shapefiletrialpoint.dbf','shapefiletrialline.dbf']
    export_file_path1 = tk.filedialog.asksaveasfilename(defaultextension='.shp',title='Save the points as')
    path = [export_file_path1]
    if os.path.exists('shapefiletrialline.shp'): 
        export_file_path2 = tk.filedialog.asksaveasfilename(defaultextension='.shp',title='Save the lines as')
        path.append(export_file_path2)
    targetDir = os.getcwd()
    
    try:
        i = 0
        for file in shp_files:
            oldshp = os.path.join(targetDir, file)
            newshp = path[i]
            os.rename(oldshp,newshp)
            i += 1
    except: ''
            
    try:
        i = 0    
        for file in shx_files:
            oldshx = os.path.join(targetDir, file)
            newshx = path[i].replace(".shp",".shx")
            os.rename(oldshx,newshx)
            i += 1
    except: ''
        
    try:
        i = 0    
        for file in dbf_files:
            olddbf = os.path.join(targetDir, file)
            newdbf = path[i].replace(".shp",".dbf")
            os.rename(olddbf,newdbf)
            i += 1
    except: ''
       
def exportdxf():
    export_file_path = tk.filedialog.asksaveasfilename(defaultextension='.dxf')
    dxf_doc.saveas(export_file_path)

def exporttxt():
    export_file_path = tk.filedialog.asksaveasfilename(defaultextension='.txt')
    with open(export_file_path,'w') as txt_report:
        print(report,file = txt_report)
        txt_report.close()
          
def disable_export():
    filesubmenu.entryconfigure('Export to excel',state='disabled')
    filesubmenu.entryconfigure('Export to csv',state='disabled')
    filesubmenu.entryconfigure('Export to shapefile',state='disabled')
    filesubmenu.entryconfigure('Export to dxf',state='disabled')
    filesubmenu.entryconfigure('Export to txt',state='disabled')
        
def traverseforget(): frametraverse0.place_forget()
def levelforget(): framelevel0.place_forget()
def hydroforget(): framehydro0.place_forget()
def areaforget(): framearea0.place_forget()
def resectionforget(): frameresection0.place_forget()
def intersectionforget(): frameintersection0.place_forget()

def homepage():
    master.title('Survey Companion')
    traverseforget()
    levelforget()
    hydroforget()
    areaforget()
    resectionforget()
    intersectionforget()
    disable_export()

def traverse():
    master.title('Survey Companion: Traverse Computation')
    #traverseforget()
    levelforget()
    hydroforget()
    areaforget()
    resectionforget()
    intersectionforget()
    disable_export()
    
    frametraverse0.place(x=1,y=1,relheight=1.0,relwidth=1.0)
    #frametraverse00.grid(row=0,column=0,sticky=NSEW)
    frametraverse1.grid(row=0,column=0,sticky='nw')
    frametraverse5.grid(row=2,column=0,sticky='nw')
    frametraverse4.grid(row=1,column=0,columnspan=3,sticky='nw')     
    
def level():
    master.title('Survey Companion: Level Reduction and Adjustment')
    traverseforget()
    #levelforget()
    hydroforget()
    areaforget()
    resectionforget()
    intersectionforget()
    disable_export()
    
    framelevel1.grid(row=0,column=0,sticky='nw')
    framelevel0.place(x=1,y=1,relheight=1.0,relwidth=1.0)
    #framelevel00.grid(row=0,column=0,sticky=NW)
    framelevel4.grid(row=1,column=0,columnspan=3,sticky='nw')
    framelevel5.grid(row=2,column=0,sticky='nw')    
    
def hydro():
    master.title('Survey Companion: Sounding Data Reduction and Filtering')
    traverseforget()
    levelforget()
    #hydroforget()
    areaforget()
    resectionforget()
    intersectionforget()
    disable_export()
    
    framehydro0.place(x=1,y=1,relheight=1.0,relwidth=1.0)
    #framehydro00.grid(row=0,column=0,sticky=NW)
    framehydro1.grid(row=0,column=0,sticky='nw')
    framehydro5.grid(row=3,column=0,columnspan=2,sticky='nw')
    framehydro4.grid(row=1,column=0,columnspan=3,sticky='nw')
    framehydro6.grid(row=2,column=0,columnspan=3,sticky='nw')    
    labelhydro7.grid(row=5,column=0,columnspan=2,sticky='w')
    open_tide_file.grid(row=6,column=0,columnspan=3,sticky='e')

def area1():
    master.title('Survey Companion: Area Computation')
    traverseforget()
    levelforget()
    hydroforget()
    #areaforget()
    resectionforget()
    intersectionforget()
    disable_export()
    
    framearea0.place(x=1,y=1,relheight=1.0,relwidth=1.0)
    framearea2.grid_forget()
    plot1.grid_forget()
    computearea1.grid_forget()
    framearea4.grid(row=1,column=0,columnspan=3, sticky='nw')
    plot.grid()
    computearea.grid()
    
def area2():
    master.title('Survey Companion: Area Computation')
    traverseforget()
    levelforget()
    hydroforget()
    #areaforget()
    resectionforget()
    intersectionforget()
    disable_export()
    
    framearea0.place(x=1,y=1,relheight=1.0,relwidth=1.0)
    framearea4.grid_forget()
    plot.grid_forget()
    computearea.grid_forget()
    framearea2.grid(row=0,column=1, sticky = 'nw')
    plot1.grid()
    computearea1.grid()
   
def resection():
    master.title('Survey Companion: Three Point Resection')
    traverseforget()
    levelforget()
    hydroforget()
    areaforget()
    #resectionforget()
    intersectionforget()
    disable_export()
    
    frameresection0.place(x=1,y=1,relheight=1.0,relwidth=1.0)
    frameresection1.grid(row=0,column=0,sticky='nw')
    frameresection2.grid(row=0,column=1,sticky = 'nw')
    
def intersection1_forget():
    frameintersection0.place(x=1,y=1,relheight=1.0,relwidth=1.0)
    computeintersection1.grid_forget()
    labelintersection9.grid_forget()
    labelintersection10.grid_forget()
    labelintersection9unit.grid_forget()
    labelintersection10unit.grid_forget()
    Azimuth_AO_zz_intersection.grid_forget()
    Azimuth_BO_zz_intersection.grid_forget()
    labelintersection8.grid_forget()
    frameintersection5.grid_forget()

def intersection2_forget():
    frameintersection0.place(x=1,y=1,relheight=1.0,relwidth=1.0)
    computeintersection2.grid_forget()
    labelintersection91.grid_forget()
    labelintersection101.grid_forget()
    labelintersection91unit.grid_forget()
    labelintersection101unit.grid_forget()
    Azimuth_AO_ad_intersection.grid_forget()
    Distance_BO_ad_intersection.grid_forget()
    labelintersection81.grid_forget()
    frameintersection51.grid_forget()

def intersection3_forget():
    frameintersection0.place(x=1,y=1,relheight=1.0,relwidth=1.0)
    computeintersection3.grid_forget()
    labelintersection92.grid_forget()
    labelintersection102.grid_forget()
    labelintersection92unit.grid_forget()
    labelintersection102unit.grid_forget()
    Distance_AO_dd_intersection.grid_forget()
    Distance_BO_dd_intersection.grid_forget()
    labelintersection82.grid_forget()
    frameintersection52.grid_forget()

def intersection4_forget():
    frameintersection0.place(x=1,y=1,relheight=1.0,relwidth=1.0)
    computeintersection4.grid_forget()
    labelintersection93.grid_forget()
    labelintersection103.grid_forget()
    labelintersection93unit.grid_forget()
    labelintersection103unit.grid_forget()
    Angle_BAO_aa_intersection.grid_forget()
    Angle_ABO_aa_intersection.grid_forget()
    labelintersection83.grid_forget()
    frameintersection53.grid_forget()
    labelintersection_order.grid_forget()
    option.grid_forget()
    
def intersection5_forget():
    frameintersection0.place(x=1,y=1,relheight=1.0,relwidth=1.0)
    computeintersection5.grid_forget()
    labelintersection94.grid_forget()
    labelintersection104.grid_forget()
    labelintersection94unit.grid_forget()
    labelintersection104unit.grid_forget()
    Angle_BAO_aa_sidesection.grid_forget()
    Angle_AOB_aa_sidesection.grid_forget()
    labelintersection84.grid_forget()
    frameintersection54.grid_forget()
    labelintersection_order.grid_forget()
    option.grid_forget()
    
def intersection6_forget():
    frameintersection0.place(x=1,y=1,relheight=1.0,relwidth=1.0)
    computeintersection6.grid_forget()
    labelintersection95.grid_forget()
    labelintersection105.grid_forget()
    labelintersection95unit.grid_forget()
    labelintersection105unit.grid_forget()
    Angle_AOB_ad_sidesection.grid_forget()
    Distance_AO_ad_sidesection.grid_forget()
    labelintersection85.grid_forget()
    frameintersection55.grid_forget()
    labelintersection_order.grid_forget()
    option.grid_forget()
   
def intersection1():
    master.title('Survey Companion: Azimuth_Azimuth Intersection')
    traverseforget()
    levelforget()
    hydroforget()
    areaforget()
    resectionforget()
    #intersectionforget()
    disable_export()
    
    intersection2_forget()
    intersection3_forget()
    intersection4_forget()
    intersection5_forget()
    intersection6_forget()
    
    computeintersection1.grid()
    labelintersection9.grid(row=6,column=0,sticky='e')
    labelintersection10.grid(row=7,column=0,sticky='e')
    Azimuth_AO_zz_intersection.grid(row=6, column=1)
    Azimuth_BO_zz_intersection.grid(row=7, column=1)
    labelintersection9unit.grid(row=6,column=2,sticky='w')
    labelintersection10unit.grid(row=7,column=2,sticky='w')
    labelintersection8.grid(row=5,column=0,columnspan=3,sticky='w')
    frameintersection5.grid(row=2,column=0,columnspan=2)
    
def intersection2():
    master.title('Survey Companion: Azimuth_Distance Intersection')
    traverseforget()
    levelforget()
    hydroforget()
    areaforget()
    resectionforget()
    #intersectionforget()
    disable_export()
    
    intersection1_forget()
    intersection3_forget()
    intersection4_forget()
    intersection5_forget()
    intersection6_forget()
    
    computeintersection2.grid()
    labelintersection91.grid(row=6,column=0,sticky='e')
    labelintersection101.grid(row=7,column=0,sticky='e')
    Azimuth_AO_ad_intersection.grid(row=6, column=1)
    Distance_BO_ad_intersection.grid(row=7, column=1)
    labelintersection91unit.grid(row=6,column=2,sticky='w')
    labelintersection101unit.grid(row=7,column=2,sticky='w')
    labelintersection81.grid(row=5,column=0,columnspan=3,sticky='w')
    frameintersection51.grid(row=2,column=0,columnspan=2)

def intersection3():
    master.title('Survey Companion: Distance_Distance Intersection')
    traverseforget()
    levelforget()
    hydroforget()
    areaforget()
    resectionforget()
    #intersectionforget()
    disable_export()
    
    intersection1_forget()
    intersection2_forget()
    intersection4_forget() 
    intersection5_forget()
    intersection6_forget()
    
    computeintersection3.grid()
    labelintersection92.grid(row=6,column=0,sticky='e')
    labelintersection102.grid(row=7,column=0,sticky='e')
    Distance_AO_dd_intersection.grid(row=6, column=1)
    Distance_BO_dd_intersection.grid(row=7, column=1)
    labelintersection92unit.grid(row=6,column=2,sticky='w')
    labelintersection102unit.grid(row=7,column=2,sticky='w')
    labelintersection82.grid(row=5,column=0,columnspan=3,sticky='w')
    frameintersection52.grid(row=2,column=0,columnspan=2)
    labelintersection_order.grid(row=8, column=0,columnspan=3,sticky='w')
    option.grid(row=9, column=2,sticky = 'e')

def intersection4():
    master.title('Survey Companion: Angle_Angle Intersection')
    traverseforget()
    levelforget()
    hydroforget()
    areaforget()
    resectionforget()
    #intersectionforget()
    disable_export()
    
    intersection1_forget()
    intersection2_forget()  
    intersection3_forget()
    intersection5_forget()
    intersection6_forget()
    
    computeintersection4.grid()
    labelintersection93.grid(row=6,column=0,sticky='e')
    labelintersection103.grid(row=7,column=0,sticky='e')
    Angle_BAO_aa_intersection.grid(row=6, column=1)
    Angle_ABO_aa_intersection.grid(row=7, column=1)
    labelintersection93unit.grid(row=6,column=2,sticky='w')
    labelintersection103unit.grid(row=7,column=2,sticky='w')
    labelintersection83.grid(row=5,column=0,columnspan=3,sticky='w')
    frameintersection53.grid(row=2,column=0,columnspan=2)
    labelintersection_order.grid(row=8, column=0,columnspan=3,sticky='w')
    option.grid(row=9, column=2,sticky = 'e')
    
def intersection5():
    master.title('Survey Companion: Angle_Angle Sidesection')
    traverseforget()
    levelforget()
    hydroforget()
    areaforget()
    resectionforget()
    #intersectionforget()
    disable_export()
    
    intersection1_forget()
    intersection2_forget()  
    intersection3_forget()
    intersection4_forget()
    intersection6_forget()
    
    computeintersection5.grid()
    labelintersection94.grid(row=6,column=0,sticky='e')
    labelintersection104.grid(row=7,column=0,sticky='e')
    Angle_BAO_aa_sidesection.grid(row=6, column=1)
    Angle_AOB_aa_sidesection.grid(row=7, column=1)
    labelintersection94unit.grid(row=6,column=2,sticky='w')
    labelintersection104unit.grid(row=7,column=2,sticky='w')
    labelintersection84.grid(row=5,column=0,columnspan=3,sticky='w')
    frameintersection54.grid(row=2,column=0,columnspan=2)
    labelintersection_order.grid(row=8, column=0,columnspan=3,sticky='w')
    option.grid(row=9, column=2,sticky = 'e')
    
def intersection6():
    master.title('Survey Companion: Angle_Distance Sidesection')
    traverseforget()
    levelforget()
    hydroforget()
    areaforget()
    resectionforget()
    #intersectionforget()
    disable_export()
    
    intersection1_forget()
    intersection2_forget()  
    intersection3_forget()
    intersection4_forget()
    intersection5_forget()
    
    computeintersection6.grid()
    labelintersection95.grid(row=6,column=0,sticky='e')
    labelintersection105.grid(row=7,column=0,sticky='e')
    Angle_AOB_ad_sidesection.grid(row=6, column=1)
    Distance_AO_ad_sidesection.grid(row=7, column=1)
    labelintersection95unit.grid(row=6,column=2,sticky='w')
    labelintersection105unit.grid(row=7,column=2,sticky='w')
    labelintersection85.grid(row=5,column=0,columnspan=3,sticky='w')
    frameintersection55.grid(row=2,column=0,columnspan=2)
    labelintersection_order.grid(row=8, column=0,columnspan=3,sticky='w')
    option.grid(row=9, column=2,sticky = 'e')
    
def speak(speech):
    engine = pyttsx3.init()
    engine.setProperty('rate',150)
    engine.say(speech)
    engine.runAndWait()    
    
def traverse_calculation():
    try:
        file = labeltraverse6['text']
        if file == 'No file Selected': 
            messagebox.showwarning("Error","Please open a data file to process")
            return
        if file.endswith('.xlsx') or file.endswith('.xls'):
            fil = pd.read_excel(file,engine = 'openpyxl')
            etf = fil.loc[0:(len(fil)-1), 'EASTINGS']
            ntf = fil.loc[0:(len(fil)-1), 'NORTHINGS']
            sta_id = fil.loc[0:(len(fil)), 'STATION ID']
            observed_angle = fil.iloc[0:(len(fil)), 2]
            dist = fil.iloc[0:(len(fil)), 3]
            stn = fil.iloc[0:(len(fil)), 0]
            stn_to = fil.iloc[0:(len(fil)), 4]
            
        else: 
            messagebox.showwarning("Error","Please your data file is in the wrong file format.\nTry again with an MS Excel data file")
            return
        
        method = traversemethod.get()
        if method == 'Choose Computation Method:':
            messagebox.showwarning("Error","Please select the method of\nTraverse Adjustment to use")
            return
        #----------------Reducing the face left and face right angles to obtain the mean angles--------------------------
        Mean_angle=[]
        increment=0
        while increment < len(observed_angle):
            firstfl=dms_to_dgs(observed_angle.iloc[increment])
            secondfl=dms_to_dgs(observed_angle.iloc[increment+1])
            firstfr=dms_to_dgs(observed_angle.iloc[increment+2])
            secondfr=dms_to_dgs(observed_angle.iloc[increment+3])
            firstdiff=secondfl-firstfl
            if firstdiff < 0: firstdiff += 360
            seconddiff=firstfr-secondfr
            if seconddiff < 0: seconddiff += 360
            mean_angle=(firstdiff+seconddiff)/2
            Mean_angle.append(mean_angle)
            increment += 4
               
        #-----------------------Computing the opening bearing of the network--------------------       
        Opening_DE = etf[1] - etf[0]
        Opening_DN = ntf[1] - ntf[0]
        firstB_B = first_bearing(Opening_DE, Opening_DN)
        if firstB_B < 180: first_back_bearing = firstB_B + 180
        else : first_back_bearing = firstB_B - 180
        #----------------------Computing the closing bearing of the network-----------------------------
        Closing_DE = etf[3] - etf[2]
        Closing_DN = ntf[3] - ntf[2]
        Closing_bearing = first_bearing(Closing_DE, Closing_DN)
        
        #------------------------Computing the backward and forward bearings of the traverse legs--------------------------
        Back_Bearings = []
        Back_Bearings.append(firstB_B) 
        Forward_Bearings = []
        increment2=0
        for angle in Mean_angle:
            forward_bearing=(angle+Back_Bearings[increment2])%360
            Forward_Bearings.append(forward_bearing)
            if forward_bearing < 180: new_back_bearing = forward_bearing + 180
            else : new_back_bearing = forward_bearing - 180
            Back_Bearings.append(new_back_bearing)
            increment2 += 1
        
        #--------------------------Computing the corresponding corrections to the bearings of the traverse legs----------------------------    
        Corrections = []
        Bearing_misclosure = Closing_bearing - Forward_Bearings[len(Forward_Bearings)-1]
        increment3 = 1
        while increment3 <= len(Forward_Bearings):
            correction = ((increment3)/(len(Forward_Bearings)))*(Bearing_misclosure)
            Corrections.append(correction)
            increment3 += 1
            
        #--------------------------Computing the corresponding corrected bearings of the traverse legs----------------------------  
        Corrected_bearing = [sum(pair) for pair in zip(Forward_Bearings,Corrections)]
        
        #---------------------------Compiling the distance list and removing empty cells in distance---------------------------
        Distance = []
        for distance in dist:
            distance = str(distance)
            if distance == "nan": distance = int("-1")
            else: distance = float(distance)
            Distance.append(distance)
            for distance in Distance:
                if distance == -1:
                    Distance.remove(distance)
        lastdist = Distance[len(Distance)-1]
        
        #--------------------------Computing the total distance of the traverse network---------------------------
        Total_distance = sum(Distance)
        if len(Distance) == len(Mean_angle):
            Total_distance = Total_distance - Distance[len(Distance)-1]
            Distance.remove(Distance[len(Distance)-1])
        
        #---------------------------Compiling the station id list and removing empty cells in station id---------------------------
        Station_ID = []
        for station_id in stn:
            station_id = str(station_id)
            if station_id == "nan": station_id = int("-1")
            else: distance = str(station_id)
            Station_ID.append(station_id)
            for station_id in Station_ID:
                if station_id == -1:
                    Station_ID.remove(station_id)
        Station_ID.remove(Station_ID[0])
        
        #---------------------------Compiling the station_to list and removing---------------------------
        Station_To = []
        incrementa = 1
        while incrementa < len(stn_to):
            station_to = stn_to[incrementa]
            Station_To.append(station_to)
            incrementa += 4
        
        #------------------------Computing and compiling the change in eastings into a list-----------------------
        Change_in_Eastings = []            
        for distance,bearing in zip(Distance,Corrected_bearing):
            Change_in_easting = easting_compt(distance,bearing)
            Change_in_Eastings.append(Change_in_easting) 
        Sum_Change_in_Eastings = sum(Change_in_Eastings)
        
        total_east = 0
        for value in Change_in_Eastings:
            total_east += abs(value)
            
        #------------------------Computing and compiling the corresponding corrections to change in eastings into a list-----------------------
        Closing_Easting = etf[2] - etf[0]
        correction_to_eastings = Sum_Change_in_Eastings - Closing_Easting
        
        if method == "Bowditch Rule":
            Correction_to_Eastings = []
            for distance in Distance:
                correction = (distance/Total_distance)*correction_to_eastings
                Correction_to_Eastings.append(correction)
        if method == "Transit Rule":
            Correction_to_Eastings = []
            for east in Change_in_Eastings:
                correction = (abs(east)/total_east)*correction_to_eastings
                Correction_to_Eastings.append(correction)
    
        #------------------------Computing and compiling the corresponding corrected change in eastings into a list-----------------------
        Corrected_Eastings = []   
        for change_in_easting,correction in zip(Change_in_Eastings,Correction_to_Eastings):
            corrected_easting = change_in_easting - correction
            Corrected_Eastings.append(corrected_easting)
        
        #------------------------Computing and compiling the change in northings into a list-----------------------
        Change_in_Northings = []            
        for distance,bearing in zip(Distance,Corrected_bearing):
            Change_in_northing = northing_compt(distance,bearing)
            Change_in_Northings.append(Change_in_northing)
        Sum_Change_in_Northings = sum(Change_in_Northings)
        
        total_north = 0
        for value in Change_in_Northings:
            total_north += abs(value)
            
        #------------------------Computing and compiling the corresponding corrections to change in northings into a list-----------------------    
        Closing_Northing = ntf[2] - ntf[0]
        correction_to_northings = Sum_Change_in_Northings - Closing_Northing
        
        if method == "Bowditch Rule":
            Correction_to_Northings = []
            for distance in Distance:
                correction = (distance/Total_distance)*correction_to_northings
                Correction_to_Northings.append(correction)
        if method == "Transit Rule":
            Correction_to_Northings = []
            for north in Change_in_Northings:
                correction = (abs(north)/total_north)*correction_to_northings
                Correction_to_Northings.append(correction)
        
        #------------------------Computing and compiling the corresponding corrected change in northings into a list-----------------------
        Corrected_Northings = []   
        for change_in_northing,correction in zip(Change_in_Northings,Correction_to_Northings):
            corrected_northing = change_in_northing - correction
            Corrected_Northings.append(corrected_northing)
        
        linear_misclosure = math.sqrt((correction_to_eastings)**2 + (correction_to_northings)**2)
        
        #-----------------------Computing the Adjusted Easting and Northing coordinates---------------------------  
        Eastings = [etf[0]]
        Northings = [ntf[0]]
        Station_IDs = [sta_id[0]]
        Easting = etf[0]
        Northing = ntf[0]
        for station_id,easting,northing in zip(Station_ID,Corrected_Eastings,Corrected_Northings):
            Easting += easting
            Northing += northing 
            Eastings.append(round(Easting,3))   
            Northings.append(round(Northing,3))
            Station_IDs.append(station_id)
        
        #----------------------Presenting the result in a tabular data frame---------------------
       
        #-------------------------Presenting the result in a graph-----------------------------
        Starting_Easting = [etf[0],etf[1]]
        Starting_Northing = [ntf[0],ntf[1]]
        Closing_Easting = [etf[2],etf[3]]   
        Closing_Northing = [ntf[2],ntf[3]]  
        
        IDs1 = [sta_id[0],sta_id[1]]
        IDs2 = [sta_id[0]]
        for i in Station_ID: IDs2.append(i)
        IDs3 = [sta_id[2],sta_id[3]]
        
        try:
            files_to_unhide = ['shapefiletrialpoint.shp','shapefiletrialpoint.shx','shapefiletrialpoint.dbf','shapefiletrialline.shp','shapefiletrialline.shx','shapefiletrialline.dbf']
            for fn in files_to_unhide:
                p = os.popen('attrib -h ' + fn)
                p.read()
                p.close() 
        except: ''
        
        with shapefile.Writer("shapefiletrialpoint") as w:
            for value in zip(Starting_Easting,Starting_Northing): w.point(value[0],value[1])
            for value in zip(Eastings,Northings): w.point(value[0],value[1])
            for value in zip(Closing_Easting,Closing_Northing): w.point(value[0],value[1])
            w.field('ID_Group')
            w.field('Station_ID','C','40')
            for value in IDs1: w.record('Start',value)
            for value in IDs2: w.record('Traverse Station',value)
            for value in IDs3: w.record('End',value)
            w.close
            
        
            
        with shapefile.Writer("shapefiletrialline") as w3:
            w3.field('Line_ID','C','40')
            w3.line([[list(xy) for xy in zip(Starting_Easting,Starting_Northing)]])
            w3.line([[list(xy) for xy in zip(Eastings,Northings)]])
            w3.line([[list(xy) for xy in zip(Closing_Easting,Closing_Northing)]])
            w3.record('Line connecting Starting Controls')
            w3.record('Line connecting Traverse Stations')
            w3.record('Line connecting Closing Controls')
            w3.close
        
        try:
            files_to_hide = ['shapefiletrialpoint.shp','shapefiletrialpoint.shx','shapefiletrialpoint.dbf','shapefiletrialline.shp','shapefiletrialline.shx','shapefiletrialline.dbf']
            for fn in files_to_hide:
                p = os.popen('attrib +h ' + fn)
                p.read()
                p.close()
        except: ''
        
        global fig1
        fig1, plot = plt.subplots(figsize = (7.5,7.5))
        plt.plot(Starting_Easting,Starting_Northing,color='red',marker='s',markerfacecolor='red',markersize=6)#,linewidth=1) 
        plt.plot(Closing_Easting,Closing_Northing,color='orange',marker='s',markerfacecolor='orange',markersize=6)#,linewidth=1)
        plt.plot(Eastings, Northings,color='blue',marker='s',markerfacecolor='blue')
        plt.axis('scaled')
        for i in range(len(Eastings)): plt.annotate(IDs2[i], (Eastings[i], Northings[i] + 0.2))
        for i in range(len(Starting_Easting)): plt.annotate(IDs1[i], (Starting_Easting[i], Starting_Northing[i] + 0.2))
        for i in range(len(Closing_Easting)): plt.annotate(IDs3[i], (Closing_Easting[i], Closing_Northing[i] + 0.2))
        
        plt.grid(True,linewidth=0.5)
        plot.set_xlabel('Eastings')
        plot.set_ylabel('Northings')
        plot.set_title('Graph of the Traverse Network')
        plot.legend(['Starting Controls','Closing Controls', 'Traverse Legs'],title='Legend')      
        
        
        if checkvartraverse.get() == 1:
            createplot('Plot of the Traverse Network')
        
        global dxf_doc
        dxf_doc = ezdxf.new(setup=True)
        msp = dxf_doc.modelspace()
        dxf_doc.layers.new(name = 'Line_connecting_Starting_Controls', dxfattribs = {'color': 1 ,'linetype':'Continuous'})
        msp.add_lwpolyline(list(xy for xy in zip(Starting_Easting,Starting_Northing)),dxfattribs={'layer':'Line_connecting_Starting_Controls'})
        dxf_doc.layers.new(name = 'Line_connecting_Traverse_Stations', dxfattribs = {'color': 5 ,'linetype':'Continuous'})
        msp.add_lwpolyline(list(xy for xy in zip(Eastings,Northings)),dxfattribs={'layer':'Line_connecting_Traverse_Stations'})
        dxf_doc.layers.new(name = 'Line_connecting_Closing_Controls', dxfattribs = {'color': 20 ,'linetype':'Continuous'})
        msp.add_lwpolyline(list(xy for xy in zip(Closing_Easting,Closing_Northing)),dxfattribs={'layer':'Line_connecting_Closing_Controls'})

        
        for ids,position in zip(IDs1,(zip(Starting_Easting,Starting_Northing))):
            msp.add_text(ids,dxfattribs = {'style':'OpenSans-Bold'}).set_pos(position,align = 'LEFT')
        for ids,position in zip(IDs2,(zip(Eastings,Northings))):
            msp.add_text(ids,dxfattribs = {'style':'OpenSans-Bold'}).set_pos(position,align = 'LEFT')
        for ids,position in zip(IDs3,(zip(Closing_Easting,Closing_Northing))):
            msp.add_text(ids,dxfattribs = {'style':'OpenSans-Bold'}).set_pos(position,align = 'LEFT')
            
        
        
        #------------------------Creating the output tables---------------------
        Back_Bearings.remove(Back_Bearings[len(Back_Bearings)-1])
        Station_ID.insert(0,stn[0])
        
        #-------------------Making the lists needed to have the same length
        if len(Distance) <= len(Mean_angle):
            Change_in_Eastings.append('-----')
            Change_in_Northings.append('-----')
            Correction_to_Eastings.append('-----')
            Correction_to_Northings.append('-----')
            Corrected_Eastings.append('-----')
            Corrected_Northings.append('-----')
            Distance.append(lastdist) 
         
        
        #---------------------------filling in the values in the computation sheet-----------------------------
        for numbers in range(len(Station_ID)):
            Station_ID[numbers] = str(Station_ID[numbers])
            Distance[numbers] = str(Distance[numbers])
            Back_Bearings[numbers] = str(dgs_to_dms(Back_Bearings[numbers]))
            Mean_angle[numbers] = str(dgs_to_dms(Mean_angle[numbers]))
            Forward_Bearings[numbers] = str(dgs_to_dms(Forward_Bearings[numbers]))
            Corrections[numbers] = str(dgs_to_dms(Corrections[numbers]))
            Corrected_bearing[numbers] = str(dgs_to_dms(Corrected_bearing[numbers]))
            if Change_in_Eastings[numbers] == '-----': Change_in_Eastings[numbers] = str(Change_in_Eastings[numbers])
            else: Change_in_Eastings[numbers] = str(rd3(Change_in_Eastings[numbers]))
            if Change_in_Northings[numbers] == '-----': Change_in_Northings[numbers] = str(Change_in_Northings[numbers])
            else: Change_in_Northings[numbers] = str(rd3(Change_in_Northings[numbers]))
            if Correction_to_Eastings[numbers] == '-----': Correction_to_Eastings[numbers] = str(Correction_to_Eastings[numbers])
            else: Correction_to_Eastings[numbers] = str(rd6(Correction_to_Eastings[numbers]))
            if Correction_to_Northings[numbers] == '-----': Correction_to_Northings[numbers] = str(Correction_to_Northings[numbers])
            else: Correction_to_Northings[numbers] = str(rd6(Correction_to_Northings[numbers]))
            if Corrected_Eastings[numbers] == '-----': Corrected_Eastings[numbers] = str(Corrected_Eastings[numbers])
            else: Corrected_Eastings[numbers] = str(rd6(Corrected_Eastings[numbers]))
            if Corrected_Northings[numbers] == '-----': Corrected_Northings[numbers] = str(Corrected_Northings[numbers])
            else: Corrected_Northings[numbers] = str(rd6(Corrected_Northings[numbers]))
            Eastings[numbers] = str(Eastings[numbers])
            Northings[numbers] = str(Northings[numbers])
            Station_To[numbers] = str(Station_To[numbers])
        
        
        frametraverse6 = LabelFrame(frametraverse00, text="**Adjusted Coordinates**", bd=5, bg="CadetBlue1", fg="green4",font=('Times', 14, 'bold', 'italic'))
        frametraverse6.grid(row=0, column=1, rowspan=100, sticky = 'nw')
        #----------------------------Creating the default output ----------------------------     
        global df4
        if checkvartraverse1.get() == 1: 
            Station_ID.insert(0,sta_id[1])
            Distance.insert(0,' ')
            Back_Bearings.insert(0,' ')
            Mean_angle.insert(0,' ')
            Forward_Bearings.insert(0,dgs_to_dms(first_back_bearing))
            Corrections.insert(0,' ')
            Corrected_bearing.insert(0,' ')
            Change_in_Eastings.insert(0,' ')
            Change_in_Northings.insert(0,' ')
            Correction_to_Eastings.insert(0,' ')
            Correction_to_Northings.insert(0,' ')
            Corrected_Eastings.insert(0,' ')
            Corrected_Northings.insert(0,' ')
            Eastings.append(etf[3])
            Northings.append(ntf[3])
            Station_To.insert(0,sta_id[0])
            
            data1 = {'Station At':Station_ID,'Distance':Distance , 'Back Bearing':Back_Bearings, "Mean Angle":Mean_angle,"Forward Bearing":Forward_Bearings,"corr. to Bearing":Corrections, "Corr. Bearing":Corrected_bearing,"ΔE(LsinƟ)":Change_in_Eastings,"ΔN(LcosƟ)":Change_in_Northings,"corr. to ΔE":Correction_to_Eastings,"corr. to ΔN":Correction_to_Northings,"Corr. ΔE":Corrected_Eastings,"Corr. ΔN":Corrected_Northings, 'Eastings':Eastings , 'Northings':Northings,"Station To":Station_To}
            df4 = DataFrame(data1, columns = ['Station At','Distance' , 'Back Bearing', "Mean Angle","Forward Bearing","corr. to Bearing", "Corr. Bearing","ΔE(LsinƟ)","ΔN(LcosƟ)","corr. to ΔE","corr. to ΔN","Corr. ΔE","Corr. ΔN",'Eastings','Northings',"Station To"])
            frametraverse6['text'] = "**Computation Sheet**"
            table = Table(frametraverse6,dataframe=df4,showtoolbar=True,showstatusbar=True)
            table.show()
        else:
            Station_To.insert(0,sta_id[1])
            Station_To.insert(1,sta_id[0])
            Eastings.insert(0,etf[1])
            Eastings.append(etf[3])
            Northings.insert(0,ntf[1])
            Northings.append(ntf[3])
               
            data2 = {'Station_IDs':Station_To , 'Eastings':Eastings , 'Northings':Northings}
            df4 = DataFrame(data2, columns = ['Station_IDs','Eastings','Northings'])
            table = Table(frametraverse6,dataframe=df4,showtoolbar=True,showstatusbar=True)
            table.show()
        
        
        labelbrmis = Label(frametraverse6, text='**Misclosure in Bearing** = ', bg='CadetBlue1',fg='green4',font=('Times', 12,'bold'))
        if Bearing_misclosure<0: labelbrmis['fg'] = 'red3'
        br_misclosure = labelbrmis['text'] + str(dgs_to_dms(Bearing_misclosure))
        labelbrmis['text'] = br_misclosure
        labelbrmis.grid(row=4,column=1,sticky='w')
        
        labelli_mis = Label(frametraverse6, text='**Linear Misclosure** = ', bg='CadetBlue1',fg='green4',font=('Times', 12,'bold'))
        if linear_misclosure<0: labelli_mis['fg'] = 'red3'
        li_misclosure = labelli_mis['text'] + str(rd3(linear_misclosure)) + 'm'
        labelli_mis['text'] = li_misclosure
        labelli_mis.grid(row=5,column=1,sticky='w')
        
        labelrel_pre = Label(frametraverse6, text='**Relative Precision** = ', bg='CadetBlue1',fg='green4',font=('Times', 12,'bold'))
        relative_precision = labelrel_pre['text'] + '1 : ' + str(round(Total_distance/linear_misclosure))
        labelrel_pre['text'] = relative_precision 
        labelrel_pre.grid(row=6,column=1,sticky='w')
        
        #print(Total_distance,linear_misclosure)
        
        filesubmenu.entryconfigure('Export to excel',state='normal')
        filesubmenu.entryconfigure('Export to csv',state='normal')
        filesubmenu.entryconfigure('Export to shapefile',state='normal')
        filesubmenu.entryconfigure('Export to dxf',state='normal')
        filesubmenu.entryconfigure('Export to txt',state='disabled')
        
        mnt,sec = divmod(abs(Bearing_misclosure)*3600,60)
        deg,mnt = divmod(mnt,60)
     
        global speech
        if Bearing_misclosure<0:speech = 'Misclosure in Bearing= '+'-'+str(int(deg))+'° '+str(int(mnt))+'minutes '+str(float('{:.2f}'.format(sec)))+'seconds'
        else:speech = 'Misclosure in Bearing= '+str(int(deg))+'° '+str(int(mnt))+'minutes '+str(float('{:.2f}'.format(sec)))+'seconds'
        global speech1
        speech1 = 'Linear Misclosure='+ str(rd3(linear_misclosure)) + 'm'
        global speech2
        speech2 = 'Relative Precision='+ '1 is to ' + str(round(Total_distance/linear_misclosure))
        speak(speech)
        speak(speech1)
        speak(speech2)
        
    except ValueError:
        messagebox.showwarning("Error","Please confirm that all fields in your data file\nare valid or check for missing values")
        return
    except KeyError:
        messagebox.showerror("Error","Please confirm that column headers for your\ncontrols in your data file match 'STATION ID',\n'EASTINGS' and 'NORTHINGS' in this order \n\nOr reconfirm your open data file")
        return
    except TypeError:
        messagebox.showwarning("Error","Please confirm that all fields in your data file\nare valid or check for missing values")
        return
    except FileNotFoundError:
        messagebox.showwarning("Error","Please select a data file to process")
        return
    except :
        messagebox.showwarning("Error","Oops!!!\nSomething went wrong\nPlease confirm your opened data file")
        return


def level_computation():
    try:
        file = labellevel6['text']
        if file == 'No file Selected': 
            messagebox.showwarning("Error","Please open a data file to process")
            return
        if file.endswith('.xlsx') or file.endswith('.xls'):
            fil = pd.read_excel(file,engine = 'openpyxl')
            staff_station = fil.iloc[0:(len(fil)), 0]
            dist = fil.iloc[0:(len(fil)), 1]
            back = fil.iloc[0:(len(fil)), 2]
            inter = fil.iloc[0:(len(fil)), 3]
            fore = fil.iloc[0:(len(fil)), 4]
            benchid = fil.loc[0:(len(fil)), 'Station ID']
            benchheight = fil.loc[0:(len(fil)), 'HEIGHT']
            
        else: 
            messagebox.showwarning("Error","Please your data file is in the wrong file format.\nTry again with an MS Excel data file")
            return
        
        staff_stationlist = []
        distlist = []
        backlist = []
        interlist = []
        forelist = []
        
        for station in staff_station : staff_stationlist.append(str(station))
        for distance in dist : distlist.append(str(distance))
        for value in back : backlist.append(str(value))
        for value in inter : interlist.append(str(value))
        for value in fore : forelist.append(str(value))
        
        
        prov_height = [benchheight[0]]
        HI = []
        Rise = ['']
        Fall = ['']
        count = 0
        while count<(len(backlist)-1) and count<(len(interlist)-1) and count<(len(forelist)-1):
            if backlist[count] !='nan' and interlist[count+1] !='nan':
                change = float(backlist[count]) - float(interlist[count+1])
            if backlist[count]!='nan' and forelist[count+1]!='nan':
                change = float(backlist[count]) - float(forelist[count+1])
            if interlist[count]!='nan' and interlist[count+1]!='nan':
                change = float(interlist[count]) - float(interlist[count+1])
            if interlist[count]!='nan' and forelist[count+1]!='nan':
                change = float(interlist[count]) - float(forelist[count+1]) 
            if change>=0:
                Rise.append(float(change))
                Fall.append('')
            else:
                Rise.append('')
                Fall.append(float(change))
            prov_height.append(prov_height[count]+change)  
            if backlist[count] !='nan' and prov_height[count] !='nan':
                HI.append(float(backlist[count]) + float(prov_height[count]))
            else: HI.append('')
            count += 1   
        
        if staff_station[len(staff_station)-1] == benchid[1]:
            misclosure = benchheight[1] - prov_height[len(prov_height)-1]
                
        backcount = back.count() 
        
        correction = []
        count = 0
        for height in prov_height:
            corr = ((fil.iloc[0:(count), 2]).count()/backcount)*misclosure
            correction.append(corr)
            count += 1              
            
        count = 0       
        corrected_height =  [sum(pair) for pair in zip(prov_height,correction)]
        for value in corrected_height:
            corrected_height[count] = round(value,3)
            count += 1
            
        count = 0
        for a,b,c,d,e in zip(staff_stationlist,distlist,backlist,interlist,forelist):
            if a =='nan':
                staff_stationlist[count] = ''
            if b =='nan':
                distlist[count] = ''
            if c =='nan':
                backlist[count] = ''
            if d =='nan':
                interlist[count] = ''
            if e =='nan':
                forelist[count] = ''
            count += 1    
        
        coordframelevel.grid(row=0, column=3, rowspan=100, sticky='nw') 
            
        for numbers in range(len(Rise)-1):
            staff_stationlist[numbers] = str(staff_stationlist[numbers])
            distlist[numbers] = str(distlist[numbers])
            backlist[numbers] = str(backlist[numbers])
            interlist[numbers] = str(interlist[numbers])
            forelist[numbers] = str(forelist[numbers])
            if HI[numbers] ==  '': HI[numbers] = str(HI[numbers])
            else: HI[numbers] = str(round(HI[numbers],3))
            if Rise[numbers] == '': Rise[numbers] = str(Rise[numbers])
            else: Rise[numbers] = str(round(Rise[numbers],3))
            if Fall[numbers] == '': Fall[numbers] = str(Fall[numbers])
            else: Fall[numbers] = str(round(abs(Fall[numbers]),3))
            prov_height[numbers] = str(round(prov_height[numbers],3))
            correction[numbers] = str(round(correction[numbers],9))
            corrected_height[numbers] = str(corrected_height[numbers])
        
        while len(HI) < len(Rise):HI.append('')
        
        global df4
        if checkvarlevel1.get() == 1:  
            data1 = {'Staff Station':staff_stationlist,'Chainages':distlist , 'Back Sight':backlist, "Intermediate Sight":interlist,"Fore Sight":forelist, "Height of Instrument(HI)":HI, "Rise":Rise, "Fall":Fall,"Reduced Height":prov_height, "Corrections":correction, "Corrected Reduced Height":corrected_height}
            df4 = DataFrame(data1, columns = ['Staff Station','Chainages' , 'Back Sight', "Intermediate Sight","Fore Sight","Height of Instrument(HI)","Rise","Fall","Reduced Height","Corrections","Corrected Reduced Height"])
            coordframelevel['text'] = "**Computation Sheet**"
            table = Table(coordframelevel,dataframe=df4,showtoolbar=True,showstatusbar=True)
            table.show()
        else:
            data2 = {'Staff Station':staff_stationlist,'Chainages':distlist , 'Back Sight':backlist, "Intermediate Sight":interlist,"Fore Sight":forelist,"Corrected Reduced Height":corrected_height}
            df4 = DataFrame(data2, columns = ['Staff Station','Chainages' , 'Back Sight', "Intermediate Sight","Fore Sight","Corrected Reduced Height"])
            coordframelevel['text'] = "**Corrected Reduced Levels**"
            table = Table(coordframelevel,dataframe=df4,showtoolbar=True,showstatusbar=True)
            table.show()
            
        labelmis = Label(coordframelevel, text='**Level Misclosure** = ', bg='CadetBlue1',fg='green4',font=('Times', 12,'bold'))
        if misclosure<0: labelmis['fg'] = 'red3'
        lvl_misclosure = labelmis['text'] + str(round(misclosure,3)) + ' m'
        labelmis['text'] = lvl_misclosure
        labelmis.grid(row=4,column=1,sticky='w') 
        
        filesubmenu.entryconfigure('Export to excel',state='normal')
        filesubmenu.entryconfigure('Export to csv',state='normal')
        filesubmenu.entryconfigure('Export to dxf',state='disabled')
        filesubmenu.entryconfigure('Export to shapefile',state='disabled')
        filesubmenu.entryconfigure('Export to txt',state='disabled')
        
        global speech
        speech = 'Level Misclosure='+ str(round(misclosure,3)) + 'm'
        speak(speech)
        
    except ValueError:
        messagebox.showwarning("Error","Please confirm that all fields in your data file\nare valid or check for missing values")
        return
    except KeyError:
        messagebox.showerror("Error","Please confirm that column headers for your\nbenchmarks in your data file match 'STATION ID'\nand 'HEIGHT' in this order \n\nOr reconfirm your open data file")
        return
    except TypeError:
        messagebox.showwarning("Error","Please confirm that all fields in your data file\nare valid or check for missing values")
        return
    except UnboundLocalError:
        messagebox.showwarning("Error","Oops!!!\nLooks like you are missing a staff reading in your data file")
        return
    except FileNotFoundError:
        messagebox.showwarning("Error","Please open a data file to process")
        return


def bathy_reduction():
    try:
        file = labelhydro6['text']
        if file == 'No file Selected': 
            messagebox.showwarning("Error","Please open a data file to process")
            return
        if file.endswith('.csv'):    
            open_file = open(file)
            fil = open_file.read()
            files = fil.strip().split('\n')
            
        else: 
            messagebox.showwarning("Error","Please your data file is in the wrong file format.\nTry again with an comma delimited(.csv) data file")
            return
        
        eastings = []
        northings = []
        time_values = []
        depths = []
        
        for values in files:
            time,east,north,depth = values.split(",")
            if not time.isalpha():
                time_values.append(time)
            if not east.isalpha():
                eastings.append(float(east))
            if not north.isalpha():
                northings.append(float(north))
            if not depth.isalpha():
                depths.append(float(depth))
        
        greater = greater_than.get()
        if greater == '' or greater == '0.З': greater = min(depths)
        try: greater = float(greater)
        except ValueError:  
            messagebox.showwarning("Error","Please input a valid lower depth range to filter")
            return
                
        less = less_than.get()
        if less == '' or less == 'З0': less = max(depths)
        try: less = float(less)
        except ValueError:  
            messagebox.showwarning("Error","Please input a valid upper depth range to filter")
            return
                    
        i = -1
        false_depth_index = []
        for value in depths:
            i += 1
            if value <= greater or value >= less : false_depth_index.append(i)
        for i in false_depth_index:                   
            time_values[i] = 'false_depth'
            eastings[i] = 'false_depth'
            northings[i] = 'false_depth'
            depths[i] = 'false_depth'
            
        time_values = [i for i in time_values if i != 'false_depth']    
        eastings = [i for i in eastings if i != 'false_depth']    
        northings = [i for i in northings if i != 'false_depth']    
        depths = [i for i in depths if i != 'false_depth']    
        
        i = 0
        for value in time_values:
            time_values[i] = hms_to_dgs(value)
            i += 1
        
        i = 0
        for value in range(len(time_values)-1):
            if float(time_values[i+1]) < float(time_values[i]): time_values[i+1] = time_values[i+1] + 12
            i += 1
        
        tide_file = labelhydro8['text']
        if tide_file == 'No file Selected': 
            messagebox.showwarning("Error","Please open a tide data file to process")
            return
        if tide_file.endswith('.csv'):
            open_tide_file = open(tide_file)
            tide_fil = open_tide_file.read()
            tide_files = tide_fil.strip().split('\n')
        else: 
            messagebox.showwarning("Error","Please your tide data file is in the wrong file format.\nTry again with an comma delimited(.csv) tide data file")
            return
        
        tide_times = []
        tide_levels = []
        
        for values in tide_files:
            tide_time,tide_level = values.split(",")
            if not tide_time.isalpha():
                tide_times.append(tide_time)
            if not tide_level.isalpha():
                tide_levels.append(float(tide_level))
                
        i = 0
        for value in tide_times:
            tide_times[i] = hms_to_dgs(value)
            i += 1
        
        i = 0
        for value in range(len(tide_times)-1):
            if float(tide_times[i+1]) < float(tide_times[i]): tide_times[i+1] = tide_times[i+1] + 12
            i += 1
        
        Water_levels = []
        Reduced_Depths = []
        for value in time_values:
            upper_bound = min(i for i in tide_times if i > value)
            upper_bound = tide_times.index(upper_bound)
            lower_bound =  upper_bound - 1
            change_in_level = ((tide_levels[upper_bound]-tide_levels[lower_bound]) * (value-tide_times[lower_bound])) / (tide_times[upper_bound]-tide_times[lower_bound])
            water_level = change_in_level + tide_levels[lower_bound]
            Water_levels.append(water_level)
        
        for depth,level in zip(depths,Water_levels):
            reduced_depth = depth - level
            Reduced_Depths.append(reduced_depth)
        
        
        i = 0
        dist_interval = distance_interval.get()
        if dist_interval == '' or dist_interval == 'З': dist_interval = 0
        try: dist_interval = float(dist_interval)
        except ValueError:  
            messagebox.showwarning("Error","Please input a valid distance interval to filter")
            return
        
        current = time_values[i]
        current_east = eastings[0]
        current_north = northings[0]
        current_time = time_values[0]
        current_depth = depths[0]
        filtered_eastings = [current_east]
        filtered_northings = [current_north]
        filtered_times = [current_time]
        filtered_depths = [current_depth]
        while time_values.index(current)+1 < len(time_values):
            if distance2(current_east,current_north,eastings[i+1],northings[i+1]) >= dist_interval:
                current_east = eastings[i+1]
                current_north = northings[i+1]
                current_time = time_values[i+1]
                current_depth = depths[i+1]
                filtered_eastings.append(current_east)
                filtered_northings.append(current_north)
                filtered_times.append(current_time)
                filtered_depths.append(current_depth)
                current = time_values[i+1]
                i += 1  
            else: 
                current = time_values[i+1]
                i += 1
                
        framehydro7.grid(row=0, column=3,sticky = 'nw',rowspan=100)
            
        data2 = {'Eastings':filtered_eastings, 'Northings':filtered_northings, 'Depths':filtered_depths}
        df4 = DataFrame(data2, columns = ['Eastings','Northings','Depths'])
        table = Table(framehydro7,dataframe=df4,showtoolbar=True,showstatusbar=True)
        table.show()
        
        try:
            files_to_unhide = ['shapefiletrialpoint.shp','shapefiletrialpoint.shx','shapefiletrialpoint.dbf','shapefiletrialline.shp','shapefiletrialline.shx','shapefiletrialline.dbf']
            for fn in files_to_unhide:
                p = os.popen('attrib -h ' + fn)
                p.read()
                p.close() 
        except: ''
        
        with shapefile.Writer(r"shapefiletrialpoint") as w:
            for value in zip(filtered_eastings,filtered_northings):w.point(value[0],value[1])
            w.field('Eastings')
            w.field('Northings')
            w.field('Depths')
            for value in zip(filtered_eastings,filtered_northings,filtered_depths): w.record(value[0],value[1],value[2])
            w.close
            
        lineshp = ['shapefiletrialline.shp','shapefiletrialline.shx','shapefiletrialline.dbf']
        try: 
            for file in lineshp: os.remove(file)
        except: ''
            
        try:
            files_to_hide = ['shapefiletrialpoint.shp','shapefiletrialpoint.shx','shapefiletrialpoint.dbf','shapefiletrialline.shp','shapefiletrialline.shx','shapefiletrialline.dbf']
            for fn in files_to_hide:
                p = os.popen('attrib +h ' + fn)
                p.read()
                p.close()
        except: ''
        
        global fig1
        fig1, plot = plt.subplots(figsize = (7.5,7.5))
        plt.scatter(filtered_eastings, filtered_northings,color='blue',marker='o', s=1)#,markersize=1)#,markerfacecolor='blue')
        plt.axis('scaled')
        for i in range(len(filtered_eastings)): plt.annotate(filtered_depths[i], (filtered_eastings[i], filtered_northings[i]))
        
        plt.grid(True,linewidth=0.5)
        plot.set_xlabel('Eastings')
        plot.set_ylabel('Northings')
        plot.set_title('Spot Depths Map')
        plot.legend(['Spot Depths'],title='Legend')      
        
        if checkvarhydro.get() == 1:
            createplot('Plot of the Spot Depths')
        
        
        global dxf_doc
        dxf_doc = ezdxf.new(setup=True)
        msp = dxf_doc.modelspace()
        dxf_doc.layers.new(name = 'Spot Depths', dxfattribs = {'color': 26 ,'linetype':'Continuous'})
        for depth,position in zip(filtered_depths,(zip(filtered_eastings,filtered_northings))):
            msp.add_text(depth,dxfattribs = {'layer':'Spot Depths','style':'OpenSans-Bold'}).set_pos(position,align = 'MIDDLE_CENTER')

        
        filesubmenu.entryconfigure('Export to excel',state='normal')
        filesubmenu.entryconfigure('Export to csv',state='normal')
        filesubmenu.entryconfigure('Export to shapefile',state='normal')
        filesubmenu.entryconfigure('Export to dxf',state='normal')
        filesubmenu.entryconfigure('Export to txt',state='normal')
        
        global speech
        speech = 'You filtered off'+str(int((len(eastings)-len(filtered_eastings))))+'fixes.'
        global speech1
        speech1 = 'You have'+str(int(len(filtered_eastings)))+'fixes left.' 
        speak(speech)
        speak(speech1)
        
    except UnicodeDecodeError:
        messagebox.showerror("Error","Please confirm that you opened the correct \ndata files or if they are in .csv file extension")
        return
    except ValueError:
        messagebox.showwarning("Error","Please confirm that all fields in your data files\nare valid or check for missing values")
        return
    except FileNotFoundError:
        messagebox.showwarning("Error","Seems you are yet to open either\nyour data file or your tide data file")
        return


def Area_Computation():
    try:
        labelareaoutput.configure(state = 'normal')
        labelareaoutput.delete('1.0','end')
        global fig1
        def gettext():
            coordinates = entry.get('1.0', 'end') 
            return coordinates
        if len(gettext()) > 1:
            file = gettext()
            files = file.strip().split('\n')
            
            forward=0
            backward=0
            try:
                firstx=float(files[0].split(",")[0])
                firsty=float(files[0].split(",")[1])
                oldx=float(files[0].split(",")[0])
                oldy=float(files[0].split(",")[1])
            except ValueError:
                messagebox.showwarning("Error",'There is an invalid coordinate entry in\nline 1 of your coordinate entry')
                return
            except IndexError: 
                messagebox.showerror("Error",'There is an incomplete coordinate pair in\nline 1 of your coordinate entry')
                return
                    
            if len(files) < 3: 
                messagebox.showerror("Error","Area cannot be bounded by less than 3 lines")
                return
                                
            area_ans = 'The area bounded by the coordinates: \n'
            X_values = []
            Y_values = []
            Perimeter = 0
            count=1
            for coordinate in files:
                if coordinate == "close" or coordinate == "c":
                    break
                else:
                    try:
                        newx=float(coordinate.split(",")[0])
                        newy=float(coordinate.split(",")[1])
                    except ValueError:
                        messagebox.showwarning("Error","There is an invalid coordinate entry in\nline "+str(count)+' of your coordinate entry')
                        return
                    except IndexError: 
                        messagebox.showerror("Error","There is an incomplete coordinate pair in\nline "+str(count)+' of your coordinate entry')
                        return
                    X_values.append(newx)
                    Y_values.append(newy)
                    area_ans = area_ans + str(newx) + ' , ' + str(newy) + '\n'
                    forward=forward+(oldx*newy)
                    backward=backward+(oldy*newx)
                    Perimeter += distance2(oldx,oldy,newx,newy)
                    oldx=newx
                    oldy=newy
                    count+=1
                    
            
            Perimeter += distance2(oldx,oldy,firstx,firsty)
            X_values.append(firstx)
            Y_values.append(firsty)
            forward=forward+(oldx*firsty)
            backward=backward+(oldy*firstx)
            Area=(forward-backward)/2
            Area=abs(Area)
            area_ans = area_ans + " **Equals " +  str(Area) + " sqrm\n"
            area_ans = area_ans + " **Perimeter equals " + str(rd3(Perimeter)) + " m\n"
            
            labelareaoutput.insert(tk.END,area_ans)
            labelareaoutput.configure(state = 'disable')
            entry.delete('1.0','end')
            
           
            try:
                files_to_unhide = ['shapefiletrialpoint.shp','shapefiletrialpoint.shx','shapefiletrialpoint.dbf','shapefiletrialline.shp','shapefiletrialline.shx','shapefiletrialline.dbf']
                for fn in files_to_unhide:
                    p = os.popen('attrib -h ' + fn)
                    p.read()
                    p.close() 
            except: ''
            
            with shapefile.Writer(r"shapefiletrialpoint") as w:
                for value in zip(X_values,Y_values): w.point(value[0],value[1])
                w.field('ID_Group')
                w.field('Eastings')
                w.field('Northings')
                for value in zip(X_values,Y_values): w.record('vertex',value[0],value[1])
                w.close
            
            with shapefile.Writer(r"shapefiletrialline") as w3:
                w3.field('Line_ID','C','40')
                w3.poly([[list(xy) for xy in zip(X_values,Y_values)]])
                w3.record('Line connecting region vertices')
                w3.close
            
            try:
                files_to_hide = ['shapefiletrialpoint.shp','shapefiletrialpoint.shx','shapefiletrialpoint.dbf','shapefiletrialline.shp','shapefiletrialline.shx','shapefiletrialline.dbf']
                for fn in files_to_hide:
                    p = os.popen('attrib +h ' + fn)
                    p.read()
                    p.close()
            except: ''
            
            global speech
            speech = 'The area of the region='+str(Area) + "square meters."
            global speech1
            speech1 = 'The perimeter of the region='+str(rd3(Perimeter)) + "m"
            
        else:
            file = labelarea6['text'] 
            if file == 'No file Selected': 
                messagebox.showwarning("Error","Please open a data file to process")
                return
            directory,file_extension = file.split('.')
            filee = directory +'.'+ file_extension
            if file_extension == 'txt' or file_extension == 'csv':
                open_filee = open(filee)
                fil = open_filee.read()
                files = fil.strip().split('\n')
                if len(files) < 3: 
                    messagebox.showerror("Error","Area cannot be bounded by less than 3 lines")
                    return
                 
                forward=0
                backward=0
                try:
                    firstx=float(files[0].split(",")[0])
                    firsty=float(files[0].split(",")[1])
                    oldx=float(files[0].split(",")[0])
                    oldy=float(files[0].split(",")[1])
                except ValueError:
                    messagebox.showwarning("Error","There is an invalid coordinate entry in\nline 1 of your coordinate file")
                    return
                except IndexError: 
                    messagebox.showerror("Error","There is an incomplete coordinate pair in\nline 1 of your coordinate file")
                    return
                area_ans = 'The area bounded by the coordinates: \n'
                
                X_values = []
                Y_values = []
                Perimeter = 0
                count=1
                for coordinate in files:
                    if coordinate == "close" or coordinate == "c":
                        break
                    else:
                        try:
                            newx=float(coordinate.split(",")[0])
                            newy=float(coordinate.split(",")[1])
                        except ValueError:
                            messagebox.showwarning("Error","There is an invalid coordinate entry in\nline "+str(count)+' of your coordinate file')
                            return
                        except IndexError: 
                            messagebox.showerror("Error","There is an incomplete coordinate pair in\nline "+str(count)+' of your coordinate file')
                            return
                        X_values.append(newx)
                        Y_values.append(newy)
                        area_ans = area_ans + str(newx) + ' , ' + str(newy) + '\n'
                        forward=forward+(oldx*newy)
                        backward=backward+(oldy*newx)
                        Perimeter += distance2(oldx,oldy,newx,newy)
                        oldx=newx
                        oldy=newy
                        count+=1
                Perimeter += distance2(oldx,oldy,firstx,firsty)       
                X_values.append(firstx)
                Y_values.append(firsty)
                forward=forward+(oldx*firsty)
                backward=backward+(oldy*firstx)
                Area=(forward-backward)/2
                Area=abs(Area)
                area_ans = area_ans + " **Equals " +  str(Area) + " sqrm\n"
                area_ans = area_ans + " **Perimeter equals " + str(rd3(Perimeter)) + " m\n"
                labelareaoutput.insert(tk.END,area_ans)
                labelareaoutput.configure(state = 'disable')
                
                #global speech
                speech = 'The area of the region='+str(Area) + "square meters."
                #global speech1
                speech1 = 'The perimeter of the region='+str(rd3(Perimeter)) + "m"
                        
            elif file_extension == 'xlsx' or file_extension == 'xls':
                fil = pd.read_excel(filee,engine = 'openpyxl')
                x = list(fil.iloc[0:(len(fil)-1), 0])
                y = list(fil.iloc[0:(len(fil)-1), 1])
                
                X = []
                Y = []
                Perimeter = 0
                forward=0
                backward=0
                
                area_ans = 'The area bounded by the coordinates: \n'
                
                for xvalue,yvalue in zip(x,y):
                    xvalue = str(xvalue)
                    yvalue = str(yvalue)
                    if xvalue == "nan" and yvalue == "nan" : 
                        xvalue = str("empty")
                        yvalue = str("empty")
                    else: 
                        xvalue = str(xvalue)
                        yvalue = str(yvalue)
                    X.append(xvalue)
                    Y.append(yvalue)
                    for xvalue,yvalue in zip(X,Y):
                        if xvalue == "empty":X.remove(xvalue) 
                        if yvalue == "empty":Y.remove(yvalue)     
                          

                if len(X) < 3 or len(Y) < 3: 
                    messagebox.showerror("Error","Area cannot be bounded by less than 3 lines")
                    return
                
                count=1
                for xvalue,yvalue in zip(X,Y):
                    xvalue = str(xvalue)
                    yvalue = str(yvalue)
                    if xvalue=="nan" or yvalue=="nan":
                        messagebox.showerror("Error","There is an incomplete coordinate pair in\nline "+str(count)+' of your coordinate file')
                        return
                    count+=1
                
                try:
                    firstx=float(X[0])
                    firsty=float(Y[0])
                    oldx=float(X[0])
                    oldy=float(Y[0]) 
                except ValueError:
                    messagebox.showwarning("Error","There is an invalid coordinate entry in\nline 1 of your coordinate file")
                    return
                count=0 
                
                for easting,northing in zip(X,Y):
                    try:
                        newx=float(X[count])
                        newy=float(Y[count]) 
                        area_ans = area_ans + str(newx) + ' , ' + str(newy) + '\n'
                        forward=forward+(oldx*newy)
                        backward=backward+(oldy*newx)
                        Perimeter += distance2(oldx,oldy,newx,newy)
                        oldx=newx
                        oldy=newy
                        count += 1
                    except ValueError:
                        messagebox.showwarning("Error","There is an invalid coordinate entry in\nline "+str(count+1)+' of your coordinate file')
                        return
                    
                
                Perimeter += distance2(oldx,oldy,firstx,firsty)
                forward=forward+(oldx*firsty)
                backward=backward+(oldy*firstx)
                Area=(forward-backward)/2
                Area=abs(Area)
                area_ans = area_ans + " **Equals " +  str(Area) + " sqrm\n"
                area_ans = area_ans + " **Perimeter equals " + str(rd3(Perimeter)) + " m\n"
                
               
                labelareaoutput.insert(tk.END,area_ans)
                labelareaoutput.configure(state = 'disable')
                X.append(firstx)
                Y.append(firsty)
                
                X_values = X
                Y_values = Y
                count = 0
                for x,y in zip(X,Y):
                    X_values[count] = X[count] = float(x)
                    Y_values[count] = Y[count] = float(y)
                    count += 1
                    
                #global speech
                speech = 'The area of the region='+str(Area) + "square meters."
                #global speech1
                speech1 = 'The perimeter of the region='+str(rd3(Perimeter)) + "m"
                
            else: 
                messagebox.showwarning("Error","Please your data file is in the wrong file format.\nTry again with an accepted data file format")
                return
                     
            try:
                files_to_unhide = ['shapefiletrialpoint.shp','shapefiletrialpoint.shx','shapefiletrialpoint.dbf','shapefiletrialline.shp','shapefiletrialline.shx','shapefiletrialline.dbf']
                for fn in files_to_unhide:
                    p = os.popen('attrib -h ' + fn)
                    p.read()
                    p.close() 
            except: ''
            
            with shapefile.Writer(r"shapefiletrialpoint") as w:
                for value in zip(X_values,Y_values): w.point(value[0],value[1])
                w.field('ID_Group')
                w.field('Eastings')
                w.field('Northings')
                for value in zip(X_values,Y_values): w.record('vertex',value[0],value[1])
                w.close
            
            with shapefile.Writer(r"shapefiletrialline") as w3:
                w3.field('Line_ID','C','40')
                w3.poly([[list(xy) for xy in zip(X_values,Y_values)]])
                w3.record('Line connecting region vertices')
                w3.close            
            
            try:
                files_to_hide = ['shapefiletrialpoint.shp','shapefiletrialpoint.shx','shapefiletrialpoint.dbf','shapefiletrialline.shp','shapefiletrialline.shx','shapefiletrialline.dbf']
                for fn in files_to_hide:
                    p = os.popen('attrib +h ' + fn)
                    p.read()
                    p.close()
            except: ''
        
        frameareaoutput.grid(row=0, column=2,sticky = 'nw',rowspan=5)
        labelareaoutput.grid(row=0, column=0,sticky = 'nw')
        
        if checkvararea.get() == 1:
            global fig1
            fig1, plot = plt.subplots(figsize = (7.5,7.5))
            plt.plot(X_values, Y_values,color='red',marker='s',markerfacecolor='blue')
            plt.scatter(X_values, Y_values,color='blue',marker='s',s=1)#,linewidth=1)
            plt.axis('scaled')
            
            plt.grid(True,linewidth=0.5)
            plot.set_xlabel('Eastings')
            plot.set_ylabel('Northings')
            plot.set_title('Area of the Region')
            plot.legend(['Bounding lines of the region','Vertices'],title='Legend')      
            createplot('Plot of the Enclosed Area')
            
      
        global report
        report = area_ans
        global dxf_doc
        dxf_doc = ezdxf.new(setup=True)
        msp = dxf_doc.modelspace()
        dxf_doc.layers.new(name = 'Line_connecting_Region_Vertices', dxfattribs = {'color': 1 ,'linetype':'Continuous'})
        msp.add_lwpolyline(list(xy for xy in zip(X_values,Y_values)),dxfattribs={'layer':'Line_connecting_Region_Vertices'})
        
        filesubmenu.entryconfigure('Export to excel',state='disabled')
        filesubmenu.entryconfigure('Export to csv',state='disabled')
        filesubmenu.entryconfigure('Export to shapefile',state='normal')
        filesubmenu.entryconfigure('Export to dxf',state='normal')
        filesubmenu.entryconfigure('Export to txt',state='normal')
        
        speak(speech)
        speak(speech1)

    except ValueError:
        messagebox.showwarning("Error","Please enter valid coordinate values\nOr check for missing fields")
        return
    except IndexError:
        messagebox.showwarning("Error","Please confirm that every x_coordinate\nhas a corresponding y_coordinate")
        return
    except FileNotFoundError:
        messagebox.showwarning("Error","Please select a data file to process")
        return
    except FileNotFoundError:
        messagebox.showwarning("Error","Please select a data file to process")
        return


def three_point_resection():
    try:
        labelresectionoutput.configure(state = 'normal')
        labelresectionoutput.delete('1.0','end')
    
        point_a = Control_A.get()
        point_b = Control_B.get()
        point_c = Control_C.get()
        aob = Angle_AÔB.get()
        boc = Angle_BÔC.get()
    
        try: ax1,ay = point_a.split(",")
        except ValueError: 
            messagebox.showwarning("Error","There is an incomplete coordinate pair for\nControl_A in your coordinate entry")
            return
    
        try: ax1,ay = float(ax1),float(ay)
        except ValueError:  
            messagebox.showwarning("Error","Please input a valid coordinate value for Control_A")
            return
    
        try: bx,by = point_b.split(",")
        except ValueError: 
            messagebox.showwarning("Error","There is an incomplete coordinate pair for\nControl_B in your coordinate entry")
            return
    
        try: bx,by = float(bx),float(by)
        except ValueError:  
            messagebox.showwarning("Error","Please input a valid coordinate value for Control_B")
            return
    
        try: cx,cy = point_c.split(",")
        except ValueError: 
            messagebox.showwarning("Error","There is an incomplete coordinate pair for\nControl_C in your coordinate entry")
            return
    
        try: cx,cy = float(cx),float(cy)
        except ValueError:  
            messagebox.showwarning("Error","Please input a valid coordinate value for Control_C")
            return
    
        try: AÔB = dms_to_dgs(aob)
        except ValueError:  
            messagebox.showwarning("Error","Please, angular observations should be numeric\nand in the form dd mm ss eg. 25 06 14")
            return
    
        try: BÔC = dms_to_dgs(boc)
        except ValueError:  
            messagebox.showwarning("Error","Please, angular observations should be numeric\nand in the form dd mm ss eg. 25 06 14")
            return
            
        AB = distance3(point_a,point_b)
        BC = distance3(point_b,point_c)
        DE_AB = float(bx)-float(ax1)
        DN_AB = float(by)-float(ay)
        az_AB = first_bearing(DE_AB,DN_AB)
        DE_BC = float(cx)-float(bx)
        DN_BC = float(cy)-float(by)
        az_BC = first_bearing(DE_BC,DN_BC)
        
        α = 180 - (az_BC - az_AB)
        SumA_C = 360 - (α + AÔB + BÔC)
        A = math.atan((BC * math.sin(math.radians(AÔB)) * math.sin(math.radians(SumA_C)))/((AB * math.sin(math.radians(BÔC)))+(BC * math.sin(math.radians(AÔB)) * math.cos(math.radians(SumA_C)))))
        #C = math.atan((AB * math.sin(math.radians(BÔC)) * math.sin(math.radians(SumA_C)))/((BC * math.sin(math.radians(AÔB)))+(AB * math.sin(math.radians(BÔC)) * math.cos(math.radians(SumA_C)))))

        
        α1 = 180 - math.degrees(A) - AÔB
        AO = (AB * math.sin(math.radians(α1)))/(math.sin(math.radians(AÔB)))
        az_AO = az_AB + (math.degrees(A))

        
        DE_AO = AO * math.sin(math.radians(az_AO))
        DN_AO = AO * math.cos(math.radians(az_AO))
        
        ox = float(ax1) + DE_AO
        oy = float(ay) + DN_AO
        resection_ans = 'The three point resection information includes: \nControl_A = '+str(ax1)+','+str(ay)+'\nControl_B = '+str(bx)+','+str(by)+'\nControl_C = '+str(cx)+','+str(cy)+'\nAngle AÔB = '+ dgs_to_dms(AÔB) +'\nAngle BÔC = '+ dgs_to_dms(BÔC)+'\n\nYour resected Point_O = '+str(round(ox,3))+','+str(round(oy,3))
        resection_ans += '\nAdditional Information!!!\nDistance_AO = '+ str(round(abs(AO),3))+' m\nDistance_BO = '+ str(round(abs(distance3(point_b,(str(ox)+','+str(oy)))),3))+' m\nDistance_CO = '+ str(round(abs(distance3(point_c,(str(ox)+','+str(oy)))),3))+' m'
        
        labelresectionoutput.insert(tk.END,resection_ans)
        labelresectionoutput.configure(state = 'disable')
        frameresectionoutput.grid(row=0, column=2,rowspan=2, sticky = 'nw')
        labelresectionoutput.grid(sticky='w')

        
        try:
            files_to_unhide = ['shapefiletrialpoint.shp','shapefiletrialpoint.shx','shapefiletrialpoint.dbf','shapefiletrialline.shp','shapefiletrialline.shx','shapefiletrialline.dbf']
            for fn in files_to_unhide:
                p = os.popen('attrib -h ' + fn)
                p.read()
                p.close() 
        except: ''
        
        with shapefile.Writer(r"shapefiletrialpoint") as w:
            w.point(ax1,ay)
            w.point(bx,by)
            w.point(cx,cy)
            w.point(ox,oy)
            w.field('ID')
            w.field('Eastings')
            w.field('Northings')
            w.record('Control_A',ax1,ay)
            w.record('Control_B',bx,by)
            w.record('Control_C',cx,cy)
            w.record('Point_O',ox,oy)
            w.close
            
        with shapefile.Writer(r"shapefiletrialline") as w3:
            w3.field('Line_ID','C','40')
            w3.line([[[ax1,ay],[ox,oy]]])
            w3.line([[[bx,by],[ox,oy]]])
            w3.line([[[cx,cy],[ox,oy]]])
            w3.record('Line connecting Control_A and Resected point')
            w3.record('Line connecting Control_B and Resected point')
            w3.record('Line connecting Control_C and Resected point')
            w3.close
        
        try:
            files_to_hide = ['shapefiletrialpoint.shp','shapefiletrialpoint.shx','shapefiletrialpoint.dbf','shapefiletrialline.shp','shapefiletrialline.shx','shapefiletrialline.dbf']
            for fn in files_to_hide:
                p = os.popen('attrib +h ' + fn)
                p.read()
                p.close()
        except: ''
        
        X = [ax1,bx,cx,ox]
        Y = [ay,by,cy,oy]
        label = ['Control_A','Control_B','Control_C','Point_O',]
            
        if checkvarresection.get() == 1:
            global fig1
            fig1, plot = plt.subplots(figsize = (7.5,7.5))
            plt.plot([ax1,ox],[ay,oy],color='red',marker='s',markerfacecolor='blue',markersize=6)#,linewidth=1) 
            plt.plot([bx,ox],[by,oy],color='orange',marker='s',markerfacecolor='blue',markersize=6)#,linewidth=1) 
            plt.plot([cx,ox],[cy,oy],color='green',marker='s',markerfacecolor='blue',markersize=6)#,linewidth=1) 
            plt.scatter(X,Y,color='blue',s=6)#,linewidth=1)
            
            plt.axis('scaled')
            for i in range(len(X)): plt.annotate(label[i], (X[i], Y[i] + 0.2))
            
            plt.grid(True,linewidth=0.5)
            plot.set_xlabel('Eastings')
            plot.set_ylabel('Northings')
            plot.set_title('Resected point')
            plot.legend(['Line Connecting Control_A and Resected Point','Line Connecting Control_B and Resected Point','Line Connecting Control_C and Resected Point','Stations'],title='Legend')      
            createplot('Plot of the Resected point')
            
        global report
        report = str(resection_ans)
        global dxf_doc
        dxf_doc = ezdxf.new(setup=True)
        msp = dxf_doc.modelspace()
        dxf_doc.layers.new(name = 'Line connecting Controls an Resected point', dxfattribs = {'color': 1 ,'linetype':'Continuous'})
        msp.add_lwpolyline([(ax1,ay),(ox,oy)],dxfattribs={'layer':'Line connecting Controls an Resected point'})
        msp.add_lwpolyline([(bx,by),(ox,oy)],dxfattribs={'layer':'Line connecting Controls an Resected point'})
        msp.add_lwpolyline([(cx,cy),(ox,oy)],dxfattribs={'layer':'Line connecting Controls an Resected point'})

        for ids,position in zip(['Control_A','Control_B','Control_C','Point_O'],[(ax1,ay),(bx,by),(cx,cy),(ox,oy)]):
            msp.add_text(ids,dxfattribs = {'style':'OpenSans-Bold'}).set_pos(position,align = 'LEFT')
        
        filesubmenu.entryconfigure('Export to excel',state='disabled')
        filesubmenu.entryconfigure('Export to csv',state='disabled')
        filesubmenu.entryconfigure('Export to shapefile',state='normal')
        filesubmenu.entryconfigure('Export to dxf',state='normal')
        filesubmenu.entryconfigure('Export to txt',state='normal')
        
        global speech
        speech = 'Your resected Point O = '+str(round(ox,3))+'comma'+str(round(oy,3))
        speak(speech)

    except TypeError: 
        messagebox.showerror("Error","Oops!!!\nThere is no resection solution for your data\nReconfirm your data and try again")
        return


def Az_Az_Intersection():
    try:
        labelintersectionoutput.configure(state = 'normal')
        labelintersectionoutput.delete('1.0','end')
    
        point_a = Control_A_intersection.get()
        point_b = Control_B_intersection.get()
        az_ao = Azimuth_AO_zz_intersection.get()
        az_bo = Azimuth_BO_zz_intersection.get()
        
        try: ax1,ay = point_a.split(",")
        except ValueError: 
            messagebox.showwarning("Error","There is an incomplete coordinate pair for\nControl_A in your coordinate entry")
            return
    
        try: ax1,ay = float(ax1),float(ay)
        except ValueError:  
            messagebox.showwarning("Error","Please input a valid coordinate value for Control_A")
            return
    
        try: bx,by = point_b.split(",")
        except ValueError: 
            messagebox.showwarning("Error","There is an incomplete coordinate pair for\nControl_B in your coordinate entry")
            return
    
        try: bx,by = float(bx),float(by)
        except ValueError:  
            messagebox.showwarning("Error","Please input a valid coordinate value for Control_B")
            return
    
        try: az_AO = dms_to_dgs(az_ao)
        except ValueError:  
            messagebox.showwarning("Error","Please, azimuth values should be numeric\nand in the form dd mm ss eg. 25 06 14")
            return
    
        try: az_BO = dms_to_dgs(az_bo)
        except ValueError:  
            messagebox.showwarning("Error","Please, azimuth values should be numeric\nand in the form dd mm ss eg. 25 06 14")
            return
    
        if az_AO == az_BO:
            messagebox.showwarning("Error","Two parallel lines cannot intersect")
            return
        
        AB = distance3(point_a,point_b)
        DE_AB = float(bx)-float(ax1)
        DN_AB = float(by)-float(ay)
        az_AB = first_bearing(DE_AB,DN_AB)
        
        if az_AB > 180: az_BA = az_AB - 180
        else: az_BA = az_AB + 180
   
        
        Â = az_AO - az_AB
        Ḃ = az_BA - az_BO
        Ô = 180 - Â - Ḃ
        
        AO = AB * math.sin(math.radians(Ḃ)) / math.sin(math.radians(Ô))
        ox = float(ax1) + (AO * math.sin(math.radians(az_AO)))
        oy = float(ay) + (AO * math.cos(math.radians(az_AO)))
        
        intersection_ans = 'The Intersection information includes: \nControl_A = '+str(ax1)+','+str(ay)+'\nControl_B = '+str(bx)+','+str(by)+'\nAzimuth AO = '+ dgs_to_dms(az_AO) +'\nAzimuth BO = '+ dgs_to_dms(az_BO)+'\n\nYour intersected Point_O = '+str(round(ox,3))+','+str(round(oy,3))
        intersection_ans += '\n\nAdditional Information!!!\n  Distance_AO = '+ str(round(abs(AO),3))+'m\n  Distance_BO = '+ str(round(abs(distance3(point_b,(str(ox)+','+str(oy)))),3))+'m\n  Angle AOB = '+ dgs_to_dms(abs(Ô))+'\n  Angle BAO = '+ dgs_to_dms(abs(Â))+'\n  Angle ABO = '+ dgs_to_dms(abs(Ḃ))
        
        labelintersectionoutput.insert(tk.END,intersection_ans)
        labelintersectionoutput.configure(state = 'disable')
        frameintersectionoutput.grid(row=0, column=2, rowspan=2, sticky = 'nw')
        labelintersectionoutput.grid(sticky='w')
    
   
        try:
            files_to_unhide = ['shapefiletrialpoint.shp','shapefiletrialpoint.shx','shapefiletrialpoint.dbf','shapefiletrialline.shp','shapefiletrialline.shx','shapefiletrialline.dbf']
            for fn in files_to_unhide:
                p = os.popen('attrib -h ' + fn)
                p.read()
                p.close() 
        except: ''
        
        with shapefile.Writer(r"shapefiletrialpoint") as w:
            w.point(ax1,ay)
            w.point(bx,by)
            w.point(ox,oy)
            w.field('ID')
            w.field('Eastings')
            w.field('Northings')
            w.record('Control_A',ax1,ay)
            w.record('Control_B',bx,by)
            w.record('Point_O',ox,oy)
            w.close
            
        with shapefile.Writer(r"shapefiletrialline") as w3:
            w3.field('Line_ID','C','40')
            w3.line([[[ax1,ay],[ox,oy]]])
            w3.line([[[bx,by],[ox,oy]]])
            w3.record('Line connecting Control_A and Intersected point')
            w3.record('Line connecting Control_B and Intersected point')
            w3.close
        
        try:
            files_to_hide = ['shapefiletrialpoint.shp','shapefiletrialpoint.shx','shapefiletrialpoint.dbf','shapefiletrialline.shp','shapefiletrialline.shx','shapefiletrialline.dbf']
            for fn in files_to_hide:
                p = os.popen('attrib +h ' + fn)
                p.read()
                p.close()
        except: ''
        
        
        global report
        report = str(intersection_ans)
        global dxf_doc
        dxf_doc = ezdxf.new(setup=True)
        msp = dxf_doc.modelspace()
        dxf_doc.layers.new(name = 'Line connecting Controls and Intersected point', dxfattribs = {'color': 1 ,'linetype':'Continuous'})
        msp.add_lwpolyline([(ax1,ay),(ox,oy)],dxfattribs={'layer':'Line connecting Controls and Intersected point'})
        msp.add_lwpolyline([(bx,by),(ox,oy)],dxfattribs={'layer':'Line connecting Controls and Intersected point'})

        for ids,position in zip(['Control_A','Control_B','Point_O'],[(ax1,ay),(bx,by),(ox,oy)]):
            msp.add_text(ids,dxfattribs = {'style':'OpenSans-Bold'}).set_pos(position,align = 'LEFT')

        X = [ax1,bx,ox]
        Y = [ay,by,oy]
        label = ['Control_A','Control_B','Point_O',]
        
        if checkvarintersection.get() == 1:
            global fig1
            #plt.rc('axes', facecolor='yellow', edgecolor='red')
            fig1, plot = plt.subplots(figsize = (7.5,7.5))
            plt.plot([ax1,ox],[ay,oy],color='red',marker='s',markerfacecolor='blue',markersize=6)#,linewidth=1) 
            plt.plot([bx,ox],[by,oy],color='orange',marker='s',markerfacecolor='blue',markersize=6)#,linewidth=1) 
            plt.scatter(X,Y,color='blue',s=6)#,linewidth=1)
            
            plt.axis('scaled')
            for i in range(len(X)): plt.annotate(label[i], (X[i], Y[i] + 0.2))
            
            plt.grid(True,linewidth=0.5)
            plot.set_xlabel('Eastings')
            plot.set_ylabel('Northings')
            plot.set_title('Intersected point')
            plot.legend(['Line Connecting Control_A and Intersected Point','Line Connecting Control_B and Intersected Point','Stations'],title='Legend')      
            createplot('Plot of the Intersected point')
            
        
        filesubmenu.entryconfigure('Export to excel',state='disabled')
        filesubmenu.entryconfigure('Export to csv',state='disabled')
        filesubmenu.entryconfigure('Export to shapefile',state='normal')
        filesubmenu.entryconfigure('Export to dxf',state='normal')
        filesubmenu.entryconfigure('Export to txt',state='normal')
        
        global speech
        speech = 'Your intersected Point O = '+str(round(ox,3))+'comma'+str(round(oy,3))
        speak(speech)
        
    except ValueError: 
        messagebox.showerror("Error","Oops!!!\nThere is no intersection solution for your data\nReconfirm your data and try again")
        return
    except TypeError: 
        messagebox.showerror("Error","Oops!!!\nLooks like there is no intersection solution for your data\nReconfirm your data and try again")
        return

            
def Az_Di_Intersection():
    try:
        labelintersectionoutput.configure(state = 'normal')
        labelintersectionoutput.delete('1.0','end')
    
        point_a = Control_A_intersection.get()
        point_b = Control_B_intersection.get()
        az_ao = Azimuth_AO_ad_intersection.get()
        bo = Distance_BO_ad_intersection.get()
        
        try: ax1,ay = point_a.split(",")
        except ValueError: 
            messagebox.showwarning("Error","There is an incomplete coordinate pair for\nControl_A in your coordinate entry")
            return
    
        try: ax1,ay = float(ax1),float(ay)
        except ValueError:  
            messagebox.showwarning("Error","Please input a valid coordinate value for Control_A")
            return
    
        try: bx,by = point_b.split(",")
        except ValueError: 
            messagebox.showwarning("Error","There is an incomplete coordinate pair for\nControl_B in your coordinate entry")
            return
    
        try: bx,by = float(bx),float(by)
        except ValueError:  
            messagebox.showwarning("Error","Please input a valid coordinate value for Control_B")
            return

        try: az_AO = dms_to_dgs(az_ao)
        except ValueError:  
            messagebox.showwarning("Error","Please, azimuth values should be numeric\nand in the form dd mm ss eg. 25 06 14")
            return
    
        try: BO = float(bo)
        except ValueError:  
            messagebox.showwarning("Error","Please input a valid distance value for line BO")
            return
        
        AB = distance3(point_a,point_b)
        DE_AB = float(bx)-float(ax1)
        DN_AB = float(by)-float(ay)
        az_AB = first_bearing(DE_AB,DN_AB)
        
        Â = az_AO - az_AB
        
        if (((2*AB*math.cos(math.radians(Â)))**2) - (4*(AB**2-BO**2))) < 0:
            messagebox.showwarning("Error","The distance BO is too short and cannot intersect line AO")
            return
        
        AO_1 = ((2*AB*math.cos(math.radians(Â))) + (math.sqrt(((2*AB*math.cos(math.radians(Â)))**2) - (4*(AB**2-BO**2)))))/2
        AO_2 = ((2*AB*math.cos(math.radians(Â))) - (math.sqrt(((2*AB*math.cos(math.radians(Â)))**2) - (4*(AB**2-BO**2)))))/2
        
        Ô1 = math.degrees(math.acos((AO_1**2 + BO**2 - AB**2)/(2*BO*AO_1)))
        Ô2 = math.degrees(math.acos((AO_2**2 + BO**2 - AB**2)/(2*BO*AO_2)))
            
        #BAO1 = math.degrees(math.acos((AB**2 + AO_1**2 - BO**2)/(2*AB*AO_1)))
        ABO1 = math.degrees(math.acos((AB**2 + BO**2 - AO_1**2)/(2*AB*BO)))
        #BAO2 = math.degrees(math.acos((AB**2 + AO_2**2 - BO**2)/(2*AB*AO_2)))
        ABO2 = math.degrees(math.acos((AB**2 + BO**2 - AO_2**2)/(2*AB*BO)))
        
        ox_1 = float(ax1) + (AO_1 * math.sin(math.radians(az_AO)))
        oy_1 = float(ay) + (AO_1 * math.cos(math.radians(az_AO)))
        ox_2 = float(ax1) + (AO_2 * math.sin(math.radians(az_AO)))
        oy_2 = float(ay) + (AO_2 * math.cos(math.radians(az_AO)))
        
        intersection_ans = 'The Intersection information includes: \nControl_A = '+str(ax1)+','+str(ay)+'\nControl_B = '+str(bx)+','+str(by)+'\nAzimuth AO = '+ dgs_to_dms(az_AO) +'\nDistance BO = '+ str(BO)+'m'+'\n\nYour possible intersected points include:\nPoint_O_1 = '+str(round(ox_1,3))+','+str(round(oy_1,3))+'\nPoint_O_2 = '+str(round(ox_2,3))+','+str(round(oy_2,3))
        intersection_ans += '\n\nAdditional Information!!!\n  Angle BÂO = '+ dgs_to_dms(abs(Â))+'\n**Possible Distances of AO include:\n    Distance AO_1 = '+ str(round(abs(AO_1),3))+'m\n    Distance AO_2 = '+ str(round(abs(AO_2),3))+'m'+'\n**Possible angles at B include:\n    Angle AḂO_1 = '+ dgs_to_dms(abs(ABO1))+'\n    Angle AḂO_2 = '+ dgs_to_dms(abs(ABO2))+'\n**Possible angles at O include:\n    Angle AÔB_1 = '+ dgs_to_dms(abs(Ô1))+'\n    Angle AÔB_2 = '+ dgs_to_dms(abs(Ô2))
        

        labelintersectionoutput.insert(tk.END,intersection_ans)
        labelintersectionoutput.configure(state = 'disable')
        frameintersectionoutput.grid(row=0, column=2, rowspan=2, sticky = 'nw')
        labelintersectionoutput.grid(sticky='w')
        

        try:
            files_to_unhide = ['shapefiletrialpoint.shp','shapefiletrialpoint.shx','shapefiletrialpoint.dbf','shapefiletrialline.shp','shapefiletrialline.shx','shapefiletrialline.dbf']
            for fn in files_to_unhide:
                p = os.popen('attrib -h ' + fn)
                p.read()
                p.close() 
        except: ''
        
        
        with shapefile.Writer(r"shapefiletrialpoint") as w:
            w.point(ax1,ay)
            w.point(bx,by)
            w.point(ox_1,oy_1)
            w.point(ox_2,oy_2)
            w.field('ID')
            w.field('Eastings')
            w.field('Northings')
            w.record('Control_A',ax1,ay)
            w.record('Control_B',bx,by)
            w.record('Point_O_1',ox_1,oy_1)
            w.record('Point_O_2',ox_2,oy_2)
            w.close
            
        with shapefile.Writer(r"shapefiletrialline") as w3:
            w3.field('Line_ID','C','40')
            w3.line([[[ax1,ay],[ox_1,oy_1]]])
            w3.line([[[bx,by],[ox_1,oy_1]]])
            w3.line([[[ax1,ay],[ox_2,oy_2]]])
            w3.line([[[bx,by],[ox_2,oy_2]]])
            w3.record('Line connecting Control_A and first Possible Intersected point')
            w3.record('Line connecting Control_B and first Possible Intersected point')
            w3.record('Line connecting Control_A and second Possible Intersected point')
            w3.record('Line connecting Control_B and second Possible Intersected point')
            w3.close
        
        try:
            files_to_hide = ['shapefiletrialpoint.shp','shapefiletrialpoint.shx','shapefiletrialpoint.dbf','shapefiletrialline.shp','shapefiletrialline.shx','shapefiletrialline.dbf']
            for fn in files_to_hide:
                p = os.popen('attrib +h ' + fn)
                p.read()
                p.close()
        except: ''
        
        
        global report
        intersection_ans2 = 'The Intersection information includes: \nControl_A = '+str(ax1)+','+str(ay)+'\nControl_B = '+str(bx)+','+str(by)+'\nAzimuth AO = '+ dgs_to_dms(az_AO) +'\nDistance BO = '+ str(BO)+'m'+'\n\nYour possible intersected points include:\nPoint_O_1 = '+str(round(ox_1,3))+','+str(round(oy_1,3))+'\nPoint_O_2 = '+str(round(ox_2,3))+','+str(round(oy_2,3))
        intersection_ans2 += '\n\nAdditional Information!!!\n  Angle BAO = '+ dgs_to_dms(abs(Â))+'\n**Possible Distances of AO include:\n    Distance AO_1 = '+ str(round(abs(AO_1),3))+'m\n    Distance AO_2 = '+ str(round(abs(AO_2),3))+'m'+'\n**Possible angles at B include:\n    Angle ABO_1 = '+ dgs_to_dms(abs(ABO1))+'\n    Angle ABO_2 = '+ dgs_to_dms(abs(ABO2))+'\n**Possible angles at O include:\n    Angle AOB_1 = '+ dgs_to_dms(abs(Ô1))+'\n    Angle AOB_2 = '+ dgs_to_dms(abs(Ô2))
        report = str(intersection_ans2)
        global dxf_doc
        dxf_doc = ezdxf.new(setup=True)
        msp = dxf_doc.modelspace()
        dxf_doc.layers.new(name = 'Line connecting Controls and possible Intersected point', dxfattribs = {'color': 1 ,'linetype':'Continuous'})
        msp.add_lwpolyline([(ax1,ay),(ox_1,oy_1)],dxfattribs={'layer':'Line connecting Controls and possible Intersected point'})
        msp.add_lwpolyline([(bx,by),(ox_1,oy_1)],dxfattribs={'layer':'Line connecting Controls and possible Intersected point'})
        msp.add_lwpolyline([(ax1,ay),(ox_2,oy_2)],dxfattribs={'layer':'Line connecting Controls and possible Intersected point'})
        msp.add_lwpolyline([(bx,by),(ox_2,oy_2)],dxfattribs={'layer':'Line connecting Controls and possible Intersected point'})
        
        for ids,position in zip(['Control_A','Control_B','Point_O_1','Point_O_2'],[(ax1,ay),(bx,by),(ox_1,oy_1),(ox_2,oy_2)]):
            msp.add_text(ids,dxfattribs = {'style':'OpenSans-Bold'}).set_pos(position,align = 'LEFT')

        X = [ax1,bx,ox_1,ox_2]
        Y = [ay,by,oy_1,oy_2]
        label = ['Control_A','Control_B','Point_O_1','Point_O_2']
        
        if checkvarintersection.get() == 1:
            global fig1
            fig1, plot = plt.subplots(figsize = (7.5,7.5))
            plt.plot([ax1,ox_1],[ay,oy_1],color='red',marker='s',markerfacecolor='blue',markersize=6)#,linewidth=1) 
            plt.plot([bx,ox_1],[by,oy_1],color='orange',marker='s',markerfacecolor='blue',markersize=6)#,linewidth=1) 
            plt.plot([ax1,ox_2],[ay,oy_2],color='red',marker='s',markerfacecolor='blue',markersize=6)#,linewidth=1) 
            plt.plot([bx,ox_2],[by,oy_2],color='orange',marker='s',markerfacecolor='blue',markersize=6)#,linewidth=1) 
            plt.scatter(X,Y,color='blue',s=6)#,linewidth=1)
            
            plt.axis('scaled')
            for i in range(len(X)): plt.annotate(label[i], (X[i], Y[i] + 0.2))
            
            plt.grid(True,linewidth=0.5)
            plot.set_xlabel('Eastings')
            plot.set_ylabel('Northings')
            plot.set_title('Possible Intersected points')
            plot.legend(['Line Connecting Control_A and first Possible Intersected Point','Line Connecting Control_B and first possible Intersected Point','Line Connecting Control_A and second Possible Intersected Point','Line Connecting Control_B and second possible Intersected Point','Stations'],title='Legend')      
            createplot('Plot of the Intersected point')

            
        filesubmenu.entryconfigure('Export to excel',state='disabled')
        filesubmenu.entryconfigure('Export to csv',state='disabled')
        filesubmenu.entryconfigure('Export to shapefile',state='normal')
        filesubmenu.entryconfigure('Export to dxf',state='normal')
        filesubmenu.entryconfigure('Export to txt',state='normal')

        global speech
        speech = 'Your possible intersected points include:\nPoint O1 = '+str(round(ox_1,3))+'comma'+str(round(oy_1,3))+',and Point O2 = '+str(round(ox_2,3))+'comma'+str(round(oy_2,3))
        speak(speech)
        
    except ValueError: 
        messagebox.showerror("Error","Oops!!!\nThere is no intersection solution for your data\nReconfirm your data and try again")
        return
    except TypeError: 
        messagebox.showerror("Error","Oops!!!\nLooks like there is no intersection solution for your data\nReconfirm your data and try again")
        return

def Di_Di_Intersection():
    try:
        labelintersectionoutput.configure(state = 'normal')
        labelintersectionoutput.delete('1.0','end')
    
        order = ordervar.get()
        if order == 'Order':
            messagebox.showwarning("Error","Please select the clockwise order of the triangle\nformed by the controls and the unknown point")
            return
            
        point_a = Control_A_intersection.get()
        point_b = Control_B_intersection.get()
        ao = Distance_AO_dd_intersection.get()
        bo = Distance_BO_dd_intersection.get()
        
        try: ax1,ay = point_a.split(",")
        except ValueError: 
            messagebox.showwarning("Error","There is an incomplete coordinate pair for\nControl_A in your coordinate entry")
            return
    
        try: ax1,ay = float(ax1),float(ay)
        except ValueError:  
            messagebox.showwarning("Error","Please input a valid coordinate value for Control_A")
            return
    
        try: bx,by = point_b.split(",")
        except ValueError: 
            messagebox.showwarning("Error","There is an incomplete coordinate pair for\nControl_B in your coordinate entry")
            return
    
        try: bx,by = float(bx),float(by)
        except ValueError:  
            messagebox.showwarning("Error","Please input a valid coordinate value for Control_B")
            return
    
        try: AO = float(ao)
        except ValueError:  
            messagebox.showwarning("Error","Please input a valid distance value for line AO")
            return
    
        try: BO = float(bo)
        except ValueError:  
            messagebox.showwarning("Error","Please input a valid distance value for line BO")
            return
        
        AB = distance3(point_a,point_b)
        DE_AB = float(bx)-float(ax1)
        DN_AB = float(by)-float(ay)
        az_AB = first_bearing(DE_AB,DN_AB)
        
        if AB > (AO + BO):
            messagebox.showwarning("Error","The distances AO and BO are too short and cannot intersect")
            return
        
        Â = math.acos((AB**2 + AO**2 - BO**2)/(2*AB*AO))
        Ḃ = math.acos((AB**2 + BO**2 - AO**2)/(2*AB*BO))
        Ô = math.acos((AO**2 + BO**2 - AB**2)/(2*BO*AO))
        az_AO_1 = az_AB + math.degrees(Â)
        az_AO_2 = az_AB - math.degrees(Â)
        
        ox_1 = float(ax1) + (AO * math.sin(math.radians(az_AO_1)))
        oy_1 = float(ay) + (AO * math.cos(math.radians(az_AO_1)))
        ox_2 = float(ax1) + (AO * math.sin(math.radians(az_AO_2)))
        oy_2 = float(ay) + (AO * math.cos(math.radians(az_AO_2)))
        
        if order == "BOA": ox,oy = ox_1,oy_1
        if order == "AOB": ox,oy = ox_2,oy_2
        
        intersection_ans = 'The Intersection information includes: \nControl_A = '+str(ax1)+','+str(ay)+'\nControl_B = '+str(bx)+','+str(by)+'\nDistance AO = '+ str(AO)+'m'+'\nDistance BO = '+ str(BO)+'m'+'\n\nYour intersected Point_O = '+str(round(ox,3))+','+str(round(oy,3))
        intersection_ans += '\n\nAdditional Information!!!\n  Angle AOB = '+ dgs_to_dms(abs(math.degrees(Ô)))+'\n  Angle BAO = '+ dgs_to_dms(abs(math.degrees(Â)))+'\n  Angle ABO = '+ dgs_to_dms(abs(math.degrees(Ḃ)))
        
     
        labelintersectionoutput.insert(tk.END,intersection_ans)
        labelintersectionoutput.configure(state = 'disable')
        frameintersectionoutput.grid(row=0, column=2, rowspan=2,sticky = 'nw')
        labelintersectionoutput.grid(sticky='w')
    
 
        try:
            files_to_unhide = ['shapefiletrialpoint.shp','shapefiletrialpoint.shx','shapefiletrialpoint.dbf','shapefiletrialline.shp','shapefiletrialline.shx','shapefiletrialline.dbf']
            for fn in files_to_unhide:
                p = os.popen('attrib -h ' + fn)
                p.read()
                p.close() 
        except: ''
        
        with shapefile.Writer(r"shapefiletrialpoint") as w:
            
            w.point(ax1,ay)
            w.point(bx,by)
            w.point(ox,oy)
            w.field('ID')
            w.field('Eastings')
            w.field('Northings')
            w.record('Control_A',ax1,ay)
            w.record('Control_B',bx,by)
            w.record('Point_O',ox,oy)
            w.close
            
        with shapefile.Writer(r"shapefiletrialline") as w3:
            w3.field('Line_ID','C','40')
            w3.line([[[ax1,ay],[ox,oy]]])
            w3.line([[[bx,by],[ox,oy]]])
            w3.record('Line connecting Control_A and Intersected point')
            w3.record('Line connecting Control_B and Intersected point')
            w3.close        
        
        try:
            files_to_hide = ['shapefiletrialpoint.shp','shapefiletrialpoint.shx','shapefiletrialpoint.dbf','shapefiletrialline.shp','shapefiletrialline.shx','shapefiletrialline.dbf']
            for fn in files_to_hide:
                p = os.popen('attrib +h ' + fn)
                p.read()
                p.close()
        except: ''
        
        global report
        report = str(intersection_ans)    
        global dxf_doc
        dxf_doc = ezdxf.new(setup=True)
        msp = dxf_doc.modelspace()
        dxf_doc.layers.new(name = 'Line connecting Controls and Intersected point', dxfattribs = {'color': 1 ,'linetype':'Continuous'})
        msp.add_lwpolyline([(ax1,ay),(ox,oy)],dxfattribs={'layer':'Line connecting Controls and Intersected point'})
        msp.add_lwpolyline([(bx,by),(ox,oy)],dxfattribs={'layer':'Line connecting Controls and Intersected point'})
        
        for ids,position in zip(['Control_A','Control_B','Point_O'],[(ax1,ay),(bx,by),(ox,oy)]):
            msp.add_text(ids,dxfattribs = {'style':'OpenSans-Bold'}).set_pos(position,align = 'LEFT')
        
        X = [ax1,bx,ox]
        Y = [ay,by,oy]
        label = ['Control_A','Control_B','Point_O',]
        
        if checkvarintersection.get() == 1:
            global fig1
            #plt.rc('axes', facecolor='yellow', edgecolor='red')
            fig1, plot = plt.subplots(figsize = (7.5,7.5))
            plt.plot([ax1,ox],[ay,oy],color='red',marker='s',markerfacecolor='blue',markersize=6)#,linewidth=1) 
            plt.plot([bx,ox],[by,oy],color='orange',marker='s',markerfacecolor='blue',markersize=6)#,linewidth=1) 
            plt.scatter(X,Y,color='blue',s=6)#,linewidth=1)
            
            plt.axis('scaled')
            for i in range(len(X)): plt.annotate(label[i], (X[i], Y[i] + 0.2))
            
            plt.grid(True,linewidth=0.5)
            plot.set_xlabel('Eastings')
            plot.set_ylabel('Northings')
            plot.set_title('Intersected point')
            plot.legend(['Line Connecting Control_A and Intersected Point','Line Connecting Control_B and Intersected Point','Stations'],title='Legend')      
            createplot('Plot of the Intersected point')
            
        filesubmenu.entryconfigure('Export to excel',state='disabled')
        filesubmenu.entryconfigure('Export to csv',state='disabled')
        filesubmenu.entryconfigure('Export to shapefile',state='normal')
        filesubmenu.entryconfigure('Export to dxf',state='normal')
        filesubmenu.entryconfigure('Export to txt',state='normal')

        global speech
        speech = 'Your intersected Point O = '+str(round(ox,3))+'comma'+str(round(oy,3))
        speak(speech)
        
    except ValueError: 
        messagebox.showerror("Error","Oops!!!\nThere is no intersection solution for your data\nReconfirm your data and try again")
        return
    except TypeError: 
        messagebox.showerror("Error","Oops!!!\nLooks like there is no intersection solution for your data\nReconfirm your data and try again")
        return

def Aa_Aa_Intersection():
    try:
        labelintersectionoutput.configure(state = 'normal')
        labelintersectionoutput.delete('1.0','end')
        order = ordervar.get()
        if order == 'Order':
            messagebox.showwarning("Error","Please select the clockwise order of the triangle\nformed by the controls and the unknown point")
            return
        
        if order == "BOA":
            point_a = Control_A_intersection.get()
            point_b = Control_B_intersection.get()
            bao = Angle_BAO_aa_intersection.get()
            abo = Angle_ABO_aa_intersection.get()
            
        if order == "AOB":
            point_b = Control_A_intersection.get()
            point_a = Control_B_intersection.get()
            abo = Angle_BAO_aa_intersection.get()
            bao = Angle_ABO_aa_intersection.get()
        
        try: ax1,ay = point_a.split(",")
        except ValueError: 
            messagebox.showwarning("Error","There is an incomplete coordinate pair for\nControl_A in your coordinate entry")
            return
    
        try: ax1,ay = float(ax1),float(ay)
        except ValueError:  
            messagebox.showwarning("Error","Please input a valid coordinate value for Control_A")
            return
    
        try: bx,by = point_b.split(",")
        except ValueError: 
            messagebox.showwarning("Error","There is an incomplete coordinate pair for\nControl_B in your coordinate entry")
            return
    
        try: bx,by = float(bx),float(by)
        except ValueError:  
            messagebox.showwarning("Error","Please input a valid coordinate value for Control_B")
            return
    
   
        try: BAO = dms_to_dgs(bao)
        except ValueError:  
            messagebox.showwarning("Error","Please, angular observations should be numeric\nand in the form dd mm ss eg. 25 06 14")
            return
    
        try: ABO = dms_to_dgs(abo)
        except ValueError:  
            messagebox.showwarning("Error","Please, angular observations should be numeric\nand in the form dd mm ss eg. 25 06 14")
            return
        
        AB = distance3(point_a,point_b)
        DE_AB = float(bx)-float(ax1)
        DN_AB = float(by)-float(ay)
        az_AB = first_bearing(DE_AB,DN_AB)
        
        #if az_AB > 180: az_BA = az_AB - 180
        #else: az_BA = az_AB + 180
        
        az_AO = az_AB + BAO
        #az_BO = az_BA - ABO
        
        AOB = 180 - ABO - BAO
        AO = AB*math.sin(math.radians(ABO))/math.sin(math.radians(ABO+BAO))
        BO = AB*math.sin(math.radians(BAO))/math.sin(math.radians(ABO+BAO))
        
        ox = float(ax1) + (AO * math.sin(math.radians(az_AO)))
        oy = float(ay) + (AO * math.cos(math.radians(az_AO)))
        
        if order == "BOA": AO,BO=AO,BO
        if order == "AOB": AO,BO=BO,AO
            
        intersection_ans = 'The Intersection information includes: \nControl_A = '+Control_A_intersection.get()+'\nControl_B = '+Control_B_intersection.get()+'\nAngle BÂO = '+ dgs_to_dms(dms_to_dgs(Angle_BAO_aa_intersection.get())) +'\nAngle AḂO = '+ dgs_to_dms(dms_to_dgs(Angle_ABO_aa_intersection.get())) +'\n\nYour intersected Point_O = '+str(round(ox,3))+','+str(round(oy,3))
        intersection_ans += '\n\nAdditional Information!!!\n  Distance of AO = '+ str(round(AO,3))+'m\n  Distance of BO = '+ str(round(BO,3))+'\n  Angle AÔB = '+ dgs_to_dms(abs(AOB))

        labelintersectionoutput.insert(tk.END,intersection_ans)
        labelintersectionoutput.configure(state = 'disable')
        frameintersectionoutput.grid(row=0, column=2, rowspan=2,sticky = 'nw')
        labelintersectionoutput.grid(sticky='w')
        
        
        try:
            files_to_unhide = ['shapefiletrialpoint.shp','shapefiletrialpoint.shx','shapefiletrialpoint.dbf','shapefiletrialline.shp','shapefiletrialline.shx','shapefiletrialline.dbf']
            for fn in files_to_unhide:
                p = os.popen('attrib -h ' + fn)
                p.read()
                p.close() 
        except: ''
        
        with shapefile.Writer(r"shapefiletrialpoint") as w:
            w.point(ax1,ay)
            w.point(bx,by)
            w.point(ox,oy)
            w.field('ID')
            w.field('Eastings')
            w.field('Northings')
            w.record('Control_A',ax1,ay)
            w.record('Control_B',bx,by)
            w.record('Point_O',ox,oy)
            w.close
            
        with shapefile.Writer(r"shapefiletrialline") as w3:
            w3.field('Line_ID','C','40')
            w3.line([[[ax1,ay],[ox,oy]]])
            w3.line([[[bx,by],[ox,oy]]])
            w3.record('Line connecting Control_A and Intersected point')
            w3.record('Line connecting Control_B and Intersected point')
            w3.close
        
        try:
            files_to_hide = ['shapefiletrialpoint.shp','shapefiletrialpoint.shx','shapefiletrialpoint.dbf','shapefiletrialline.shp','shapefiletrialline.shx','shapefiletrialline.dbf']
            for fn in files_to_hide:
                p = os.popen('attrib +h ' + fn)
                p.read()
                p.close()
        except: ''
        
        global report
        intersection_ans2 = 'The Intersection information includes: \nControl_A = '+Control_A_intersection.get()+'\nControl_B = '+Control_B_intersection.get()+'\nAngle BAO = '+ dgs_to_dms(dms_to_dgs(Angle_BAO_aa_intersection.get())) +'\nAngle ABO = '+ dgs_to_dms(dms_to_dgs(Angle_ABO_aa_intersection.get())) +'\n\nYour intersected Point_O = '+str(round(ox,3))+','+str(round(oy,3))
        intersection_ans2 += '\n\nAdditional Information!!!\n  Distance of AO = '+ str(round(AO,3))+'m\n  Distance of BO = '+ str(round(BO,3))+'\n  Angle AOB = '+ dgs_to_dms(abs(AOB))
        report = str(intersection_ans2)
        global dxf_doc
        dxf_doc = ezdxf.new(setup=True)
        msp = dxf_doc.modelspace()
        dxf_doc.layers.new(name = 'Line connecting Controls and Intersected point', dxfattribs = {'color': 1 ,'linetype':'Continuous'})
        msp.add_lwpolyline([(ax1,ay),(ox,oy)],dxfattribs={'layer':'Line connecting Controls and Intersected point'})
        msp.add_lwpolyline([(bx,by),(ox,oy)],dxfattribs={'layer':'Line connecting Controls and Intersected point'})
        
        if order == "BOA":
            for ids,position in zip(['Control_A','Control_B','Point_O'],[(ax1,ay),(bx,by),(ox,oy)]):
                msp.add_text(ids,dxfattribs = {'style':'OpenSans-Bold'}).set_pos(position,align = 'LEFT')
        if order == "AOB":
            for ids,position in zip(['Control_B','Control_A','Point_O'],[(ax1,ay),(bx,by),(ox,oy)]):
                msp.add_text(ids,dxfattribs = {'style':'OpenSans-Bold'}).set_pos(position,align = 'LEFT')

        
        X = [ax1,bx,ox]
        Y = [ay,by,oy]
        if order == "BOA": label = ['Control_A','Control_B','Point_O',]
        if order == "AOB": label = ['Control_B','Control_A','Point_O',]
        
        if checkvarintersection.get() == 1:
            global fig1
            fig1, plot = plt.subplots(figsize = (7.5,7.5))
            plt.plot([ax1,ox],[ay,oy],color='red',marker='s',markerfacecolor='blue',markersize=6)#,linewidth=1) 
            plt.plot([bx,ox],[by,oy],color='orange',marker='s',markerfacecolor='blue',markersize=6)#,linewidth=1) 
            plt.scatter(X,Y,color='blue',s=6)#,linewidth=1)
            
            plt.axis('scaled')
            for i in range(len(X)): plt.annotate(label[i], (X[i], Y[i] + 0.2))
            
            plt.grid(True,linewidth=0.5)
            plot.set_xlabel('Eastings')
            plot.set_ylabel('Northings')
            plot.set_title('Intersected point')
            if order == "BOA": plot.legend(['Line Connecting Control_A and Intersected Point','Line Connecting Control_B and Intersected Point','Stations'],title='Legend')      
            if order == "AOB": plot.legend(['Line Connecting Control_B and Intersected Point','Line Connecting Control_A and Intersected Point','Stations'],title='Legend')      
            createplot('Plot of the Intersected point')
        
         
        filesubmenu.entryconfigure('Export to excel',state='disabled')
        filesubmenu.entryconfigure('Export to csv',state='disabled')
        filesubmenu.entryconfigure('Export to shapefile',state='normal')
        filesubmenu.entryconfigure('Export to dxf',state='normal')
        filesubmenu.entryconfigure('Export to txt',state='normal')

        global speech
        speech = 'Your intersected Point O = '+str(round(ox,3))+'comma'+str(round(oy,3))
        speak(speech)
        
    except ValueError: 
        messagebox.showerror("Error","Oops!!!\nThere is no intersection solution for your data\nReconfirm your data and try again")
        return
    except TypeError: 
        messagebox.showerror("Error","Oops!!!\nLooks like there is no intersection solution for your data\nReconfirm your data and try again")
        return

            
def Aa_Aa_Sidesection():
    try:
        labelintersectionoutput.configure(state = 'normal')
        labelintersectionoutput.delete('1.0','end')
        order = ordervar.get()
        if order == 'Order':
            messagebox.showwarning("Error","Please select the clockwise order of the triangle\nformed by the controls and the unknown point")  
            return
        
        bao = Angle_BAO_aa_sidesection.get()
        aob = Angle_AOB_aa_sidesection.get()
        try: BAO_ = dms_to_dgs(bao)
        except ValueError:  
            messagebox.showwarning("Error","Please, angular observations should be numeric\nand in the form dd mm ss eg. 25 06 14")
            return
    
        try: AOB = dms_to_dgs(aob)
        except ValueError:  
            messagebox.showwarning("Error","Please, angular observations should be numeric\nand in the form dd mm ss eg. 25 06 14")
            return
        
        ABO_ = 180 - BAO_ - AOB
        
        if order == "BOA":
            point_a = Control_A_intersection.get()
            point_b = Control_B_intersection.get()
            BAO = BAO_
            ABO = ABO_
            
        if order == "AOB":
            point_b = Control_A_intersection.get()
            point_a = Control_B_intersection.get()
            ABO = BAO_
            BAO = ABO_
        
        try: ax1,ay = point_a.split(",")
        except ValueError: 
            messagebox.showwarning("Error","There is an incomplete coordinate pair for\nControl_A in your coordinate entry")
            return
    
        try: ax1,ay = float(ax1),float(ay)
        except ValueError:  
            messagebox.showwarning("Error","Please input a valid coordinate value for Control_A")
            return
    
        try: bx,by = point_b.split(",")
        except ValueError: 
            messagebox.showwarning("Error","There is an incomplete coordinate pair for\nControl_B in your coordinate entry")
            return
    
        try: bx,by = float(bx),float(by)
        except ValueError:  
            messagebox.showwarning("Error","Please input a valid coordinate value for Control_B")
            return
       
        AB = distance3(point_a,point_b)
        DE_AB = float(bx)-float(ax1)
        DN_AB = float(by)-float(ay)
        az_AB = first_bearing(DE_AB,DN_AB)
        
        if az_AB > 180: az_BA = az_AB - 180
        else: az_BA = az_AB + 180
        
        az_AO = az_AB + BAO
        az_BO = az_BA - ABO
        
        AO = AB*math.sin(math.radians(ABO))/math.sin(math.radians(ABO+BAO))
        BO = AB*math.sin(math.radians(BAO))/math.sin(math.radians(ABO+BAO))
        
        ox = float(ax1) + (AO * math.sin(math.radians(az_AO)))
        oy = float(ay) + (AO * math.cos(math.radians(az_AO)))
            
        if order == "BOA": AO,BO=AO,BO
        if order == "AOB": AO,BO=BO,AO
        intersection_ans = 'The Intersection information includes: \nControl_A = '+Control_A_intersection.get()+'\nControl_B = '+Control_B_intersection.get()+'\nAngle BÂO = '+ dgs_to_dms(dms_to_dgs(Angle_BAO_aa_sidesection.get())) +'\nAngle AÔB = '+ dgs_to_dms(dms_to_dgs(Angle_AOB_aa_sidesection.get())) +'\n\nYour intersected Point_O = '+str(round(ox,3))+','+str(round(oy,3))
        intersection_ans += '\n\nAdditional Information!!!\n  Distance of AO = '+ str(round(AO,3))+'m\n  Distance of BO = '+ str(round(BO,3))+'\n  Angle AḂO = '+ dgs_to_dms(abs(ABO_))

        labelintersectionoutput.insert(tk.END,intersection_ans)
        labelintersectionoutput.configure(state = 'disable')
        frameintersectionoutput.grid(row=0, column=2, rowspan=2,sticky = 'nw')
        labelintersectionoutput.grid(sticky='w')
    
    
        try:
            files_to_unhide = ['shapefiletrialpoint.shp','shapefiletrialpoint.shx','shapefiletrialpoint.dbf','shapefiletrialline.shp','shapefiletrialline.shx','shapefiletrialline.dbf']
            for fn in files_to_unhide:
                p = os.popen('attrib -h ' + fn)
                p.read()
                p.close() 
        except: ''
        
        with shapefile.Writer(r"shapefiletrialpoint") as w:
            w.point(ax1,ay)
            w.point(bx,by)
            w.point(ox,oy)
            w.field('ID')
            w.field('Eastings')
            w.field('Northings')
            w.record('Control_A',ax1,ay)
            w.record('Control_B',bx,by)
            w.record('Point_O',ox,oy)
            w.close
            
        with shapefile.Writer(r"shapefiletrialline") as w3:
            w3.field('Line_ID','C','40')
            w3.line([[[ax1,ay],[ox,oy]]])
            w3.line([[[bx,by],[ox,oy]]])
            w3.record('Line connecting Control_A and Intersected point')
            w3.record('Line connecting Control_B and Intersected point')
            w3.close        
        
        try:
            files_to_hide = ['shapefiletrialpoint.shp','shapefiletrialpoint.shx','shapefiletrialpoint.dbf','shapefiletrialline.shp','shapefiletrialline.shx','shapefiletrialline.dbf']
            for fn in files_to_hide:
                p = os.popen('attrib +h ' + fn)
                p.read()
                p.close()
        except: ''
        
        global report
        intersection_ans2 = 'The Intersection information includes: \nControl_A = '+Control_A_intersection.get()+'\nControl_B = '+Control_B_intersection.get()+'\nAngle BAO = '+ dgs_to_dms(dms_to_dgs(Angle_BAO_aa_sidesection.get())) +'\nAngle AOB = '+ dgs_to_dms(dms_to_dgs(Angle_AOB_aa_sidesection.get())) +'\n\nYour intersected Point_O = '+str(round(ox,3))+','+str(round(oy,3))
        intersection_ans2 += '\n\nAdditional Information!!!\n  Distance of AO = '+ str(round(AO,3))+'m\n  Distance of BO = '+ str(round(BO,3))+'m\n  Angle ABO = '+ dgs_to_dms(abs(ABO_))
        report = str(intersection_ans2)    
        global dxf_doc
        dxf_doc = ezdxf.new(setup=True)
        msp = dxf_doc.modelspace()
        dxf_doc.layers.new(name = 'Line connecting Controls and Intersected point', dxfattribs = {'color': 1 ,'linetype':'Continuous'})
        msp.add_lwpolyline([(ax1,ay),(ox,oy)],dxfattribs={'layer':'Line connecting Controls and Intersected point'})
        msp.add_lwpolyline([(bx,by),(ox,oy)],dxfattribs={'layer':'Line connecting Controls and Intersected point'})
        
        if order == "BOA":
            for ids,position in zip(['Control_A','Control_B','Point_O'],[(ax1,ay),(bx,by),(ox,oy)]):
                msp.add_text(ids,dxfattribs = {'style':'OpenSans-Bold'}).set_pos(position,align = 'LEFT')
        if order == "AOB":
            for ids,position in zip(['Control_B','Control_A','Point_O'],[(ax1,ay),(bx,by),(ox,oy)]):
                msp.add_text(ids,dxfattribs = {'style':'OpenSans-Bold'}).set_pos(position,align = 'LEFT')

        X = [ax1,bx,ox]
        Y = [ay,by,oy]
        if order == "BOA": label = ['Control_A','Control_B','Point_O',]
        if order == "AOB": label = ['Control_B','Control_A','Point_O',]
        
        if checkvarintersection.get() == 1:
            global fig1
            #plt.rc('axes', facecolor='yellow', edgecolor='red')
            fig1, plot = plt.subplots(figsize = (7.5,7.5))
            plt.plot([ax1,ox],[ay,oy],color='red',marker='s',markerfacecolor='blue',markersize=6)#,linewidth=1) 
            plt.plot([bx,ox],[by,oy],color='orange',marker='s',markerfacecolor='blue',markersize=6)#,linewidth=1) 
            plt.scatter(X,Y,color='blue',s=6)#,linewidth=1)
            
            plt.axis('scaled')
            for i in range(len(X)): plt.annotate(label[i], (X[i], Y[i] + 0.2))
            
            plt.grid(True,linewidth=0.5)
            plot.set_xlabel('Eastings')
            plot.set_ylabel('Northings')
            plot.set_title('Intersected point')
            if order == "BOA": plot.legend(['Line Connecting Control_A and Intersected Point','Line Connecting Control_B and Intersected Point','Stations'],title='Legend')      
            if order == "AOB": plot.legend(['Line Connecting Control_B and Intersected Point','Line Connecting Control_A and Intersected Point','Stations'],title='Legend')      
            createplot('Plot of the Intersected point')
    
        filesubmenu.entryconfigure('Export to excel',state='disabled')
        filesubmenu.entryconfigure('Export to csv',state='disabled')
        filesubmenu.entryconfigure('Export to shapefile',state='normal')
        filesubmenu.entryconfigure('Export to dxf',state='normal')
        filesubmenu.entryconfigure('Export to txt',state='normal')
        
        global speech
        speech = 'Your intersected Point O = '+str(round(ox,3))+'comma'+str(round(oy,3))
        speak(speech)
        
    except ValueError: 
        messagebox.showerror("Error","Oops!!!\nThere is no intersection solution for your data\nReconfirm your data and try again")
        return
    except TypeError: 
        messagebox.showerror("Error","Oops!!!\nLooks like there is no intersection solution for your data\nReconfirm your data and try again")
        return
    
def Aa_Di_Sidesection(): 
    try:
        labelintersectionoutput.configure(state = 'normal')
        labelintersectionoutput.delete('1.0','end')
        order = ordervar.get()
        if order == 'Order':
            messagebox.showwarning("Error","Please select the clockwise order of the triangle\nformed by the controls and the unknown point")  
            return
        
        aob = Angle_AOB_ad_sidesection.get()
        ao = Distance_AO_ad_sidesection.get()
        
        try: AOB = dms_to_dgs(aob)
        except ValueError:  
            messagebox.showwarning("Error","Please, angular observations should be numeric\nand in the form dd mm ss eg. 25 06 14")
            return
        
        try: AO = float(ao)
        except ValueError:  
            messagebox.showwarning("Error","Please input a valid distance value for line AO")
            return
    
        point_a = Control_A_intersection.get()
        point_b = Control_B_intersection.get()
        
        try: ax1,ay = point_a.split(",")
        except ValueError: 
            messagebox.showwarning("Error","There is an incomplete coordinate pair for\nControl_A in your coordinate entry")
            return
    
        try: ax1,ay = float(ax1),float(ay)
        except ValueError:  
            messagebox.showwarning("Error","Please input a valid coordinate value for Control_A")
            return
    
        try: bx,by = point_b.split(",")
        except ValueError: 
            messagebox.showwarning("Error","There is an incomplete coordinate pair for\nControl_B in your coordinate entry")
            return
    
        try: bx,by = float(bx),float(by)
        except ValueError:  
            messagebox.showwarning("Error","Please input a valid coordinate value for Control_B")
            return
    
        AB = distance3(point_a,point_b)
        DE_AB = float(bx)-float(ax1)
        DN_AB = float(by)-float(ay)
        az_AB = first_bearing(DE_AB,DN_AB)
        
        BO_1 = ((2*AO*math.cos(math.radians(AOB))) + (math.sqrt(((2*AO*math.cos(math.radians(AOB)))**2) - (4*(AO**2-AB**2)))))/2
        BO_2 = ((2*AO*math.cos(math.radians(AOB))) - (math.sqrt(((2*AO*math.cos(math.radians(AOB)))**2) - (4*(AO**2-AB**2)))))/2
                        
        BAO1 = math.degrees(math.acos((AB**2 + AO**2 - BO_1**2)/(2*AB*AO)))
        ABO1 = math.degrees(math.acos((AB**2 + BO_1**2 - AO**2)/(2*AB*BO_1)))
        BAO2 = math.degrees(math.acos((AB**2 + AO**2 - BO_2**2)/(2*AB*AO)))
        ABO2 = math.degrees(math.acos((AB**2 + BO_2**2 - AO**2)/(2*AB*BO_2)))
        
        if az_AB > 180: az_BA = az_AB - 180
        else: az_BA = az_AB + 180
        
      
        if order == "AOB":
            az_AO1 = az_AB - BAO1
            az_AO2 = az_AB - BAO2
            az_BO1 = az_BA + ABO1
            az_BO2 = az_BA + ABO1
            
        if order == "BOA":
            az_AO1 = az_AB + BAO1
            az_AO2 = az_AB + BAO2
            az_BO1 = az_BA - ABO1
            az_BO2 = az_BA - ABO2
        
        ox1 = float(ax1) + (AO * math.sin(math.radians(az_AO1)))
        oy1 = float(ay) + (AO * math.cos(math.radians(az_AO1)))
        
        ox2 = float(ax1) + (AO * math.sin(math.radians(az_AO2)))
        oy2 = float(ay) + (AO * math.cos(math.radians(az_AO2)))
    
    
        intersection_ans = 'The Intersection information includes: \nControl_A = '+Control_A_intersection.get()+'\nControl_B = '+Control_B_intersection.get()+'\nDistance AO = '+ str(AO)+'m'+'\nAngle AÔB = '+ dgs_to_dms(dms_to_dgs(Angle_AOB_ad_sidesection.get())) +'\n\n**Your possible intersected points include:\n    Point_O_1 = '+str(round(ox1,3))+','+str(round(oy1,3))+'\n    Point_O_2 = '+str(round(ox2,3))+','+str(round(oy2,3))
        intersection_ans += '\n\nAdditional Information!!!\n**Possible Distances of BO include:\n    Distance BO_1 = '+ str(round(abs(BO_1),3))+'m\n    Distance BO_2 = '+ str(round(abs(BO_2),3))+'m'+'\n**Possible angles at A include:\n    Angle BÂO_1 = '+ dgs_to_dms(abs(BAO1))+'\n    Angle BÂO_2 = '+ dgs_to_dms(abs(BAO2))+'\n**Possible angles at B include:\n    Angle AḂO_1 = '+ dgs_to_dms(abs(ABO1))+'\n    Angle AḂO_2 = '+ dgs_to_dms(abs(ABO2))
        labelintersectionoutput.insert(tk.END,intersection_ans)
        labelintersectionoutput.configure(state = 'disable')
        frameintersectionoutput.grid(row=0, column=2,rowspan=2,sticky = 'nw')
        labelintersectionoutput.grid(sticky='w')

        try:
            files_to_unhide = ['shapefiletrialpoint.shp','shapefiletrialpoint.shx','shapefiletrialpoint.dbf','shapefiletrialline.shp','shapefiletrialline.shx','shapefiletrialline.dbf']
            for fn in files_to_unhide:
                p = os.popen('attrib -h ' + fn)
                p.read()
                p.close() 
        except: ''
        
        with shapefile.Writer(r"shapefiletrialpoint") as w:
            w.point(ax1,ay)
            w.point(bx,by)
            w.point(ox1,oy1)
            w.point(ox2,oy2)
            w.field('ID')
            w.field('Eastings')
            w.field('Northings')
            w.record('Control_A',ax1,ay)
            w.record('Control_B',bx,by)
            w.record('Point_O_1',ox1,oy1)
            w.record('Point_O_2',ox2,oy2)
            w.close
            
        with shapefile.Writer(r"shapefiletrialline") as w3:
            w3.field('Line_ID','C','40')
            w3.line([[[ax1,ay],[ox1,oy1]]])
            w3.line([[[bx,by],[ox1,oy1]]])
            w3.line([[[ax1,ay],[ox2,oy2]]])
            w3.line([[[bx,by],[ox2,oy2]]])
            w3.record('Line connecting Control_A and first Possible Intersected point')
            w3.record('Line connecting Control_B and first Possible Intersected point')
            w3.record('Line connecting Control_A and second Possible Intersected point')
            w3.record('Line connecting Control_B and second Possible Intersected point')
            w3.close        
        
        try:
            files_to_hide = ['shapefiletrialpoint.shp','shapefiletrialpoint.shx','shapefiletrialpoint.dbf','shapefiletrialline.shp','shapefiletrialline.shx','shapefiletrialline.dbf']
            for fn in files_to_hide:
                p = os.popen('attrib +h ' + fn)
                p.read()
                p.close()
        except: ''
        
        intersection_ans2 = 'The Intersection information includes: \nControl_A = '+Control_A_intersection.get()+'\nControl_B = '+Control_B_intersection.get()+'\nDistance AO = '+ str(AO)+'m'+'\nAngle AOB = '+ dgs_to_dms(dms_to_dgs(Angle_AOB_ad_sidesection.get())) +'\n\n**Your possible intersected points include:\n    Point_O_1 = '+str(round(ox1,3))+','+str(round(oy1,3))+'\n    Point_O_2 = '+str(round(ox2,3))+','+str(round(oy2,3))
        intersection_ans2 += '\n\nAdditional Information!!!\n**Possible Distances of BO include:\n    Distance BO_1 = '+ str(round(abs(BO_1),3))+'m\n    Distance BO_2 = '+ str(round(abs(BO_2),3))+'m'+'\n**Possible angles at A include:\n    Angle BAO_1 = '+ dgs_to_dms(abs(BAO1))+'\n    Angle BAO_2 = '+ dgs_to_dms(abs(BAO2))+'\n**Possible angles at B include:\n    Angle ABO_1 = '+ dgs_to_dms(abs(ABO1))+'\n    Angle ABO_2 = '+ dgs_to_dms(abs(ABO2))
       
        global report
        report = str(intersection_ans2)
        global dxf_doc
        dxf_doc = ezdxf.new(setup=True)
        msp = dxf_doc.modelspace()
        dxf_doc.layers.new(name = 'Line connecting Controls and possible Intersected point', dxfattribs = {'color': 1 ,'linetype':'Continuous'})
        msp.add_lwpolyline([(ax1,ay),(ox1,oy1)],dxfattribs={'layer':'Line connecting Controls and possible Intersected point'})
        msp.add_lwpolyline([(bx,by),(ox1,oy1)],dxfattribs={'layer':'Line connecting Controls and possible Intersected point'})
        msp.add_lwpolyline([(ax1,ay),(ox2,oy2)],dxfattribs={'layer':'Line connecting Controls and possible Intersected point'})
        msp.add_lwpolyline([(bx,by),(ox2,oy2)],dxfattribs={'layer':'Line connecting Controls and possible Intersected point'})
        
        for ids,position in zip(['Control_A','Control_B','Point_O_1','Point_O_2'],[(ax1,ay),(bx,by),(ox1,oy1),(ox2,oy2)]):
            msp.add_text(ids,dxfattribs = {'style':'OpenSans-Bold'}).set_pos(position,align = 'LEFT')

        X = [ax1,bx,ox1,ox2]
        Y = [ay,by,oy1,oy2]
        label = ['Control_A','Control_B','Point_O_1','Point_O_2']
        
        if checkvarintersection.get() == 1:
            global fig1
            ##plt.rc('axes', facecolor='yellow', edgecolor='red')
            fig1, plot = plt.subplots(figsize = (7.5,7.5))
            plt.plot([ax1,ox1],[ay,oy1],color='red',marker='s',markerfacecolor='blue',markersize=6)#,linewidth=1) 
            plt.plot([bx,ox1],[by,oy1],color='orange',marker='s',markerfacecolor='blue',markersize=6)#,linewidth=1) 
            plt.plot([ax1,ox2],[ay,oy2],color='red',marker='s',markerfacecolor='blue',markersize=6)#,linewidth=1) 
            plt.plot([bx,ox2],[by,oy2],color='orange',marker='s',markerfacecolor='blue',markersize=6)#,linewidth=1) 
            plt.scatter(X,Y,color='blue',s=6)#,linewidth=1)
            
            plt.axis('scaled')
            for i in range(len(X)): plt.annotate(label[i], (X[i], Y[i] + 0.2))
            
            plt.grid(True,linewidth=0.5)
            plot.set_xlabel('Eastings')
            plot.set_ylabel('Northings')
            plot.set_title('Possible Intersected points')
            plot.legend(['Line Connecting Control_A and first Possible Intersected Point','Line Connecting Control_B and first possible Intersected Point','Line Connecting Control_A and second Possible Intersected Point','Line Connecting Control_B and second possible Intersected Point','Stations'],title='Legend')      
            createplot('Plot of the Possible Intersected point')
        
    
        filesubmenu.entryconfigure('Export to excel',state='disabled')
        filesubmenu.entryconfigure('Export to csv',state='disabled')
        filesubmenu.entryconfigure('Export to shapefile',state='normal')
        filesubmenu.entryconfigure('Export to dxf',state='normal')
        filesubmenu.entryconfigure('Export to txt',state='normal')
        
        global speech
        speech = 'Your possible intersected points include:\nPoint O1 = '+str(round(ox1,3))+'comma'+str(round(oy1,3))+',and Point O2 = '+str(round(ox2,3))+'comma'+str(round(oy2,3))
        speak(speech)
        
    except ValueError: 
        messagebox.showerror("Error","Oops!!!\nThere is no intersection solution for your data\nReconfirm your data and try again")
        return
    except TypeError: 
        messagebox.showerror("Error","Oops!!!\nLooks like there is no intersection solution for your data\nReconfirm your data and try again")
        return
    
#-----------------------Creating the button to prompt running of the traverse---------------------------------        
computetraverse = Button(frametraverse5, text='COMPUTE', command=traverse_calculation, bg='green4', fg='white', font=('verdana', 11, 'bold'),cursor='hand2')
computetraverse.grid()

computelevel = Button(framelevel5, text='REDUCE', command=level_computation, bg='green4', fg='white', font=('verdana', 11, 'bold'),cursor='hand2')
computelevel.grid()

open_tide_file = Button(framehydro2, text='OPEN', command=find2, bg='green4', fg='white', font=('verdana', 11, 'bold'),cursor='hand2')
reducesounding = Button(framehydro5, text='REDUCE', command=bathy_reduction, bg='green4', fg='white', font=('verdana', 11, 'bold'),cursor='hand2')
reducesounding.grid()

computearea = Button(framearea5, text='COMPUTE', command=Area_Computation, bg='blue', fg='white', font=('verdana', 11, 'bold'), cursor='hand2', relief='raised')
computearea1 = Button(framearea5, text='COMPUTE', command=Area_Computation, bg='blue', fg='white', font=('verdana', 11, 'bold'), cursor='hand2', relief='raised')

computeresection = Button(frameresection5, text='COMPUTE', command=three_point_resection, bg='blue', fg='white', font=('verdana', 11, 'bold'), cursor='hand2', relief='raised')
computeresection.grid()

frameintersection2.grid(row=0,column=1,sticky='nw')
computeintersection1 = Button(frameintersection5, text='COMPUTE', command=Az_Az_Intersection, bg='blue', fg='white', font=('verdana', 11, 'bold'), cursor='hand2', relief='raised')
computeintersection2 = Button(frameintersection51, text='COMPUTE', command=Az_Di_Intersection, bg='blue', fg='white', font=('verdana', 11, 'bold'), cursor='hand2', relief='raised')
computeintersection3 = Button(frameintersection52, text='COMPUTE', command=Di_Di_Intersection, bg='blue', fg='white', font=('verdana', 11, 'bold'), cursor='hand2', relief='raised')
computeintersection4 = Button(frameintersection53, text='COMPUTE', command=Aa_Aa_Intersection, bg='blue', fg='white', font=('verdana', 11, 'bold'), cursor='hand2', relief='raised')
computeintersection5 = Button(frameintersection54, text='COMPUTE', command=Aa_Aa_Sidesection, bg='blue', fg='white', font=('verdana', 11, 'bold'), cursor='hand2', relief='raised')
computeintersection6 = Button(frameintersection55, text='COMPUTE', command=Aa_Di_Sidesection, bg='blue', fg='white', font=('verdana', 11, 'bold'), cursor='hand2', relief='raised')



#============================================menu==========================================
menubar=Menu(master)
filemenu=Menu(menubar, tearoff=0)

filemenu.add_command(label="Open Data file",command=find)
filesubmenu = Menu(filemenu)
filesubmenu.add_command(label='Export to excel',command=exportexcel,state='disabled')
filesubmenu.add_command(label='Export to csv',command=exportcsv,state='disabled')
filesubmenu.add_command(label='Export to shapefile',command=exportshp,state='disabled')
filesubmenu.add_command(label='Export to dxf',command=exportdxf,state='disabled')
filesubmenu.add_command(label='Export to txt',command=exporttxt,state='disabled')
filemenu.add_cascade(label='Export output',menu=filesubmenu)
filemenu.add_separator()
filemenu.add_command(label="Back Home", command=homepage)
filemenu.add_command(label="Exit", command=Exit)
menubar.add_cascade(label="File", menu=filemenu)

traversemenu=Menu(menubar, tearoff=0)
traversemenu.add_command(label="Compute Traverse",command=traverse)
menubar.add_cascade(label='Traverse',menu=traversemenu)

levelmenu=Menu(menubar, tearoff=0)
levelmenu.add_command(label="Reduce Level",command=level)
menubar.add_cascade(label='Level',menu=levelmenu)

hydromenu=Menu(menubar, tearoff=0)
hydromenu.add_command(label="Reduce Sounding",command=hydro)
menubar.add_cascade(label='Bathymetry',menu=hydromenu)

areamenu=Menu(menubar, tearoff=0)
menubar.add_cascade(label='Area',menu=areamenu)
areasubmenu = Menu(areamenu)
areasubmenu.add_command(label='Read coordinates from file',command=area1)
areasubmenu.add_command(label='Input coordinates directly',command=area2)
areamenu.add_cascade(label='Compute Area',menu=areasubmenu)

resectionmenu=Menu(menubar, tearoff=0)
resectionmenu.add_command(label="Three Point Resection",command=resection)
menubar.add_cascade(label='Resection',menu=resectionmenu)

intersectionmenu=Menu(menubar, tearoff=0)
menubar.add_cascade(label='Intersection',menu=intersectionmenu)
intersectionsubmenu1 = Menu(intersectionmenu)
intersectionsubmenu2 = Menu(intersectionmenu)
intersectionmenu.add_cascade(label='Compute Intersection',menu=intersectionsubmenu1)
intersectionsubmenu1.add_command(label='Azimuth_Azimuth Intersection',command=intersection1)
intersectionsubmenu1.add_command(label='Azimuth_Distance Intersection',command=intersection2)
intersectionsubmenu1.add_command(label='Distance_Distance Intersection',command=intersection3)
intersectionsubmenu1.add_command(label='Angle_Angle Intersection',command=intersection4)

intersectionmenu.add_cascade(label='Compute Side section',menu=intersectionsubmenu2)
intersectionsubmenu2.add_command(label='Interior Angle Side section',command=intersection5)
intersectionsubmenu2.add_command(label='Distance Side section',command=intersection6)

helpmenu=Menu(menubar, tearoff=0)
helpmenu.add_command(label='About',command=about)
helpmenu.add_command(label='User Guide',command=User_guide)
menubar.add_cascade(label='Help',menu=helpmenu)

master.config(menu=menubar)

#-------------------Running the entire GUI-----------------------------
master.mainloop()
    
    
    
    
    
    
    
#З
#3
    
    
    
    
    
    
    
    
    
    
    
    
    