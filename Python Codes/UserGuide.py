# -*- coding: utf-8 -*-
"""
Created on Sun Dec 20 01:52:53 2020

@author: IFEANYI
"""
import tkinter as tk
from tkinter import *
from tkinter import ttk

root = Toplevel()
root.geometry('670x450')
root.title('Survey Companion: User Guide') 

icon = PhotoImage(file = 'S_Companion Icon.png')
root.iconphoto(False,icon)

panedwindow = tk.PanedWindow(root,borderwidth=10,bg='powder blue', sashwidth=5, sashrelief=GROOVE)
panedwindow.pack(fill="both", expand=True)

mainframe2 = tk.Frame(panedwindow)
frame1 = Frame(mainframe2)
f_abouthome = Frame(frame1, bg='blue')
canvasabouthome = Canvas(f_abouthome, borderwidth=0)
frameabouthome00 = Frame(canvasabouthome)

canvasabouthome.place(x=1,y=1,relheight=1.0,relwidth=1.0)
canvasabouthome.create_window((1,1), window=frameabouthome00, anchor="nw")

imageabouthome1 = PhotoImage(file='Abouthome user guide_001.png')
labelabouthome1 = Label(frameabouthome00, text='''
Designed and developed by UGWUANYI, IFEANYI ALEXANDER

    Tel.: +234 814 138 3318
    Email: ugwuanyialexifyco@gmail.com
                 ifeanyi.ugwuanyi.197580@unn.edu.ng

''',justify='left',font=('helvetica', 12,'italic','bold'))

labelabouthomeimage1 = Label(frameabouthome00, image=imageabouthome1)

labelabouthomeimage1.pack()
labelabouthome1.pack()
frame1.place(x=0,y=0,relheight=1.0,relwidth=1.0)
f_abouthome.place(x=1,y=1,relheight=1.0,relwidth=1.0)
canvasabouthome.place(x=1,y=1,relheight=1.0,relwidth=1.0)

   
def onFrameConfigure(canvas): canvas.configure(scrollregion=canvas.bbox("all")) 
def userguides(event):
    item_iid = coordinates.selection()[0]
      
    if item_iid == "I001":
        frame1 = Frame(mainframe2)
        f_abouthome = Frame(frame1, bg='blue')
        canvasabouthome = Canvas(f_abouthome, borderwidth=0)
        frameabouthome00 = Frame(canvasabouthome)
        
        canvasabouthome.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        canvasabouthome.create_window((1,1), window=frameabouthome00, anchor="nw")
    
        imageabouthome1 = PhotoImage(file='Abouthome user guide_001.png')
        labelabouthome1 = Label(frameabouthome00, text='''
    Designed and developed by UGWUANYI, IFEANYI ALEXANDER

            Tel.: +234 814 138 3318
            Email: ugwuanyialexifyco@gmail.com
                         ifeanyi.ugwuanyi.197580@unn.edu.ng

        ''',justify='left',font=('helvetica', 12,'italic','bold'))
        
        labelabouthomeimage1 = Label(frameabouthome00, image=imageabouthome1)
        
        labelabouthomeimage1.pack()
        labelabouthome1.pack()
        frame1.place(x=0,y=0,relheight=1.0,relwidth=1.0)
        f_abouthome.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        canvasabouthome.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        manageEnterWindow    
    

    elif item_iid == "I002":
        # ===========================menuhome User  Guide==================================================
        frame1 = Frame(mainframe2)
        f_menuhome = Frame(frame1, bg='blue')
        canvasmenuhome = Canvas(f_menuhome, borderwidth=0)
        framemenuhome00 = Frame(canvasmenuhome)
        
        def mousewheel(event):canvasmenuhome.yview_scroll(int(-1*(event.delta/120)),"units")
        def bind_to_wheel(event):canvasmenuhome.bind_all("<MouseWheel>",mousewheel)
        def unbind_to_wheel(event):canvasmenuhome.unbind_all("<MouseWheel>")
        f_menuhome.bind('<Enter>',bind_to_wheel)    
        f_menuhome.bind('<Leave>',unbind_to_wheel)   
        
        
        vsb = Scrollbar(f_menuhome, orient="vertical", command=canvasmenuhome.yview)
        vsb.pack(side="right", fill="y")
        hsb = Scrollbar(f_menuhome, orient="horizontal", command=canvasmenuhome.xview)
        hsb.pack(side="bottom", fill="x")
        #canvasmenuhome.pack(side="left", fill="both", expand=True)
        
        canvasmenuhome.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        
        framemenuhome00.bind("<Configure>", lambda event, canvas=canvasmenuhome: onFrameConfigure(canvasmenuhome))
        canvasmenuhome.create_window((1,1), window=framemenuhome00, anchor="nw")
        canvasmenuhome.configure(yscrollcommand=vsb.set)
        canvasmenuhome.configure(xscrollcommand=hsb.set)
        
        
        imagemenuhome1 = PhotoImage(file='menuhome user guide_001.png')
        #imagemenuhome2 = PhotoImage(file='menuhome user guide_002.png')
        imagemenuhome3 = PhotoImage(file='menuhome user guide_003.png')
        
        labelmenuhome1 = Label(framemenuhome00, text='''
        Survey Companion has the following items in its menu bar:\n
            1.	File Menu\n
            2.	Traverse Menu\n
            3.	Level Menu\n
            4.	Bathymetry Menu\n
            5.	Area Menu\n
            6.	Resection Menu\n
            7.	Intersection Menu\n
            8.	Help Menu\n
        These menus and their capabilities are further explained in the drop down.
                       
        ''',justify='left',font=('helvetica', 12,'italic'))
        
        labelmenuhomeimage1 = Label(framemenuhome00, image=imagemenuhome1)
        labelmenuhomeheading = Label(framemenuhome00, image=imagemenuhome3)
        
        labelmenuhomeheading.pack()
        labelmenuhome1.pack()
        labelmenuhomeimage1.pack()
        frame1.place(x=0,y=0,relheight=1.0,relwidth=1.0)
        f_menuhome.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        canvasmenuhome.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        manageEnterWindow    
        
    
    elif item_iid == "I003":
        # ===========================Programhome User  Guide==================================================
        frame1 = Frame(mainframe2)
        f_programhome = Frame(frame1, bg='blue')
        canvasprogramhome = Canvas(f_programhome, borderwidth=0)
        frameprogramhome00 = Frame(canvasprogramhome)
        
        def mousewheel(event):canvasprogramhome.yview_scroll(int(-1*(event.delta/120)),"units")
        def bind_to_wheel(event):canvasprogramhome.bind_all("<MouseWheel>",mousewheel)
        def unbind_to_wheel(event):canvasprogramhome.unbind_all("<MouseWheel>")
        f_programhome.bind('<Enter>',bind_to_wheel)    
        f_programhome.bind('<Leave>',unbind_to_wheel)  
        
        
        vsb = Scrollbar(f_programhome, orient="vertical", command=canvasprogramhome.yview)
        vsb.pack(side="right", fill="y")
        hsb = Scrollbar(f_programhome, orient="horizontal", command=canvasprogramhome.xview)
        hsb.pack(side="bottom", fill="x")
        #canvasprogramhome.pack(side="left", fill="both", expand=True)
        canvasprogramhome.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        
        frameprogramhome00.bind("<Configure>", lambda event, canvas=canvasprogramhome: onFrameConfigure(canvasprogramhome))
        canvasprogramhome.create_window((1,1), window=frameprogramhome00, anchor="nw")
        canvasprogramhome.configure(yscrollcommand=vsb.set)
        canvasprogramhome.configure(xscrollcommand=hsb.set)
        
        
        imageprogramhome1 = PhotoImage(file='Programhome user guide_001.png')
        #imageprogramhome2 = PhotoImage(file='Programhome user guide_002.png')
        imageprogramhome3 = PhotoImage(file='Programhome user guide_003.png')
        
        labelprogramhome1 = Label(frameprogramhome00, text='''
        Survey Companion is packed with modules that performs the following operations:
            
            1. Traverse Computation\n
            2. Level Computation\n
            3. Sounding Reduction\n
            4. Area Computation\n
            5. Three Point Resection\n
            6. Intersection\n
        These modules and their capabilities are further explained in the drop down.

        ''',justify='left',font=('helvetica', 12,'italic'))
        
        labelprogramhomeimage1 = Label(frameprogramhome00, image=imageprogramhome1)
        labelprogramhomeheading = Label(frameprogramhome00, image=imageprogramhome3)
        
        labelprogramhomeheading.pack()
        labelprogramhome1.pack()
        labelprogramhomeimage1.pack()
        frame1.place(x=0,y=0,relheight=1.0,relwidth=1.0)
        f_programhome.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        canvasprogramhome.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        manageEnterWindow    
        
    
    elif item_iid == "I004":
        # ===========================Filemenu User  Guide==================================================
        frame1 = Frame(mainframe2)
        f_filemenu = Frame(frame1, bg='blue')
        canvasfilemenu = Canvas(f_filemenu, borderwidth=0)
        framefilemenu00 = Frame(canvasfilemenu)
        
        def mousewheel(event):canvasfilemenu.yview_scroll(int(-1*(event.delta/120)),"units")
        def bind_to_wheel(event):canvasfilemenu.bind_all("<MouseWheel>",mousewheel)
        def unbind_to_wheel(event):canvasfilemenu.unbind_all("<MouseWheel>")
        f_filemenu.bind('<Enter>',bind_to_wheel)    
        f_filemenu.bind('<Leave>',unbind_to_wheel)  
        
       
        vsb = Scrollbar(f_filemenu, orient="vertical", command=canvasfilemenu.yview)
        vsb.pack(side="right", fill="y")
        hsb = Scrollbar(f_filemenu, orient="horizontal", command=canvasfilemenu.xview)
        hsb.pack(side="bottom", fill="x")
        #canvasfilemenu.pack(side="left", fill="both", expand=True)
        canvasfilemenu.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        
        framefilemenu00.bind("<Configure>", lambda event, canvas=canvasfilemenu: onFrameConfigure(canvasfilemenu))
        canvasfilemenu.create_window((1,1), window=framefilemenu00, anchor="nw")
        canvasfilemenu.configure(yscrollcommand=vsb.set)
        canvasfilemenu.configure(xscrollcommand=hsb.set)
        
        
        imagefilemenu1 = PhotoImage(file='Filemenu user guide_001.png')
        #imagefilemenu2 = PhotoImage(file='Filemenu user guide_002.png')
        imagefilemenu3 = PhotoImage(file='Filemenu user guide_003.png')
        
        labelfilemenu1 = Label(framefilemenu00, text='''
        Survey Companion "File" menu has the following functionalities and submenu:
            
        ⊞ Open Data File: This allows the user to import data files needed for the different program modules 
        
        ⊞ Export output: This submenu allows the user to export computed results to different file formats. 
        
          The “Export output” submenu has the following functions:
              
                1. Export to excel: This export functionality exports all tabular outputs to a Microsoft Excel
                   Spreadsheet (.xlsx). It can only be activated upon operations that result in tabular output. 
                   For example, traverse computation, level reduction and sounding data reduction.
                    
                2. Export to csv: This export functionality exports all tabular outputs to a comma delimited file 
                   (.csv). It can only be activated upon operations that result in tabular output. For example,  
                   traverse computation, level reduction and sounding data reduction.
                    
                3. Export to shp: This export functionality exports all plot outputs to a shapefile (.shp). It can
                   only be activated upon operations that can result in visual plots. For example,  traverse 
                   computation, sounding data reduction, area, resection and intersection computations.
                    
                4. Export to dxf: This export functionality exports all plot outputs to a drawing exchange format 
                   file (.dxf). It can only be activated upon operations that can result in visual plots. For example,
                   traverse computation, sounding data reduction, area, resection and intersection computations.
                    
                5. Export to txt: This export functionality exports all reports to a text file format (.txt). It can 
                   only be activated upon operations that can result in reports. For example, area, resection and 
                   intersection computations.
                    
        ⊞ Back Home: This allows the user to return to Survey Companion landing page. 
        
        ⊞ Exit: This allows the user to return to Survey Companion landing page.    
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        labelfilemenuimage1 = Label(framefilemenu00, image=imagefilemenu1)
        labelfilemenuheading = Label(framefilemenu00, image=imagefilemenu3)
        
        labelfilemenuheading.pack()
        labelfilemenu1.pack()
        labelfilemenuimage1.pack()
        frame1.place(x=0,y=0,relheight=1.0,relwidth=1.0)
        f_filemenu.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        canvasfilemenu.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        manageEnterWindow    
    

    elif item_iid == "I005":
        # ===========================Traversemenu User  Guide==================================================
        frame1 = Frame(mainframe2)
        f_traversemenu = Frame(frame1, bg='blue')
        canvastraversemenu = Canvas(f_traversemenu, borderwidth=0)
        frametraversemenu00 = Frame(canvastraversemenu)
        
        def mousewheel(event):canvastraversemenu.yview_scroll(int(-1*(event.delta/120)),"units")
        def bind_to_wheel(event):canvastraversemenu.bind_all("<MouseWheel>",mousewheel)
        def unbind_to_wheel(event):canvastraversemenu.unbind_all("<MouseWheel>")
        f_traversemenu.bind('<Enter>',bind_to_wheel)    
        f_traversemenu.bind('<Leave>',unbind_to_wheel)  
        
        
        vsb = Scrollbar(f_traversemenu, orient="vertical", command=canvastraversemenu.yview)
        vsb.pack(side="right", fill="y")
        hsb = Scrollbar(f_traversemenu, orient="horizontal", command=canvastraversemenu.xview)
        hsb.pack(side="bottom", fill="x")
        #canvastraversemenu.pack(side="left", fill="both", expand=True)
        canvastraversemenu.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        
        frametraversemenu00.bind("<Configure>", lambda event, canvas=canvastraversemenu: onFrameConfigure(canvastraversemenu))
        canvastraversemenu.create_window((1,1), window=frametraversemenu00, anchor="nw")
        canvastraversemenu.configure(yscrollcommand=vsb.set)
        canvastraversemenu.configure(xscrollcommand=hsb.set)
        
        
        imagetraversemenu1 = PhotoImage(file='Traversemenu user guide_001.png')
        #imagetraversemenu2 = PhotoImage(file='traversemenu user guide_002.png')
        imagetraversemenu3 = PhotoImage(file='Traversemenu user guide_003.png')
        
        labeltraversemenu1 = Label(frametraversemenu00, text='''
        Survey Companion “Traverse” menu activates the Traverse Computation Module
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        labeltraversemenuimage1 = Label(frametraversemenu00, image=imagetraversemenu1)
        labeltraversemenuheading = Label(frametraversemenu00, image=imagetraversemenu3)
        
        labeltraversemenuheading.pack()
        labeltraversemenu1.pack()
        labeltraversemenuimage1.pack()
        frame1.place(x=0,y=0,relheight=1.0,relwidth=1.0)
        f_traversemenu.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        canvastraversemenu.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        manageEnterWindow  
        
        
    elif item_iid == "I014":
        # ===========================Levelmenu User  Guide==================================================
        frame1 = Frame(mainframe2)
        f_levelmenu = Frame(frame1, bg='blue')
        canvaslevelmenu = Canvas(f_levelmenu, borderwidth=0)
        framelevelmenu00 = Frame(canvaslevelmenu)
        
        def mousewheel(event):canvaslevelmenu.yview_scroll(int(-1*(event.delta/120)),"units")
        def bind_to_wheel(event):canvaslevelmenu.bind_all("<MouseWheel>",mousewheel)
        def unbind_to_wheel(event):canvaslevelmenu.unbind_all("<MouseWheel>")
        f_levelmenu.bind('<Enter>',bind_to_wheel)    
        f_levelmenu.bind('<Leave>',unbind_to_wheel)  
        
        
        vsb = Scrollbar(f_levelmenu, orient="vertical", command=canvaslevelmenu.yview)
        vsb.pack(side="right", fill="y")
        hsb = Scrollbar(f_levelmenu, orient="horizontal", command=canvaslevelmenu.xview)
        hsb.pack(side="bottom", fill="x")
        #canvaslevelmenu.pack(side="left", fill="both", expand=True)
        canvaslevelmenu.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        
        framelevelmenu00.bind("<Configure>", lambda event, canvas=canvaslevelmenu: onFrameConfigure(canvaslevelmenu))
        canvaslevelmenu.create_window((1,1), window=framelevelmenu00, anchor="nw")
        canvaslevelmenu.configure(yscrollcommand=vsb.set)
        canvaslevelmenu.configure(xscrollcommand=hsb.set)
        
        
        imagelevelmenu1 = PhotoImage(file='Levelmenu user guide_001.png')
        #imagelevelmenu2 = PhotoImage(file='Levelmenu user guide_002.png')
        imagelevelmenu3 = PhotoImage(file='Levelmenu user guide_003.png')
        
        labellevelmenu1 = Label(framelevelmenu00, text='''
        Survey Companion “Level” menu activates the Level Computation Module
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        labellevelmenuimage1 = Label(framelevelmenu00, image=imagelevelmenu1)
        labellevelmenuheading = Label(framelevelmenu00, image=imagelevelmenu3)
        
        labellevelmenuheading.pack()
        labellevelmenu1.pack()
        labellevelmenuimage1.pack()
        frame1.place(x=0,y=0,relheight=1.0,relwidth=1.0)
        f_levelmenu.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        canvaslevelmenu.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        manageEnterWindow  
        
        
    elif item_iid == "I015":
        # ===========================Hydromenu User  Guide==================================================
        frame1 = Frame(mainframe2)
        f_hydromenu = Frame(frame1, bg='blue')
        canvashydromenu = Canvas(f_hydromenu, borderwidth=0)
        framehydromenu00 = Frame(canvashydromenu)
        
        def mousewheel(event):canvashydromenu.yview_scroll(int(-1*(event.delta/120)),"units")
        def bind_to_wheel(event):canvashydromenu.bind_all("<MouseWheel>",mousewheel)
        def unbind_to_wheel(event):canvashydromenu.unbind_all("<MouseWheel>")
        f_hydromenu.bind('<Enter>',bind_to_wheel)    
        f_hydromenu.bind('<Leave>',unbind_to_wheel)  
        
        
        vsb = Scrollbar(f_hydromenu, orient="vertical", command=canvashydromenu.yview)
        vsb.pack(side="right", fill="y")
        hsb = Scrollbar(f_hydromenu, orient="horizontal", command=canvashydromenu.xview)
        hsb.pack(side="bottom", fill="x")
        #canvashydromenu.pack(side="left", fill="both", expand=True)
        canvashydromenu.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        
        framehydromenu00.bind("<Configure>", lambda event, canvas=canvashydromenu: onFrameConfigure(canvashydromenu))
        canvashydromenu.create_window((1,1), window=framehydromenu00, anchor="nw")
        canvashydromenu.configure(yscrollcommand=vsb.set)
        canvashydromenu.configure(xscrollcommand=hsb.set)
        
        
        imagehydromenu1 = PhotoImage(file='Hydromenu user guide_001.png')
        #imagehydromenu2 = PhotoImage(file='Hydromenu user guide_002.png')
        imagehydromenu3 = PhotoImage(file='Hydromenu user guide_003.png')
        
        labelhydromenu1 = Label(framehydromenu00, text='''
        Survey Companion “Bathymetry” menu activates the Sounding Reduction Module
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        labelhydromenuimage1 = Label(framehydromenu00, image=imagehydromenu1)
        labelhydromenuheading = Label(framehydromenu00, image=imagehydromenu3)
        
        labelhydromenuheading.pack()
        labelhydromenu1.pack()
        labelhydromenuimage1.pack()
        frame1.place(x=0,y=0,relheight=1.0,relwidth=1.0)
        f_hydromenu.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        canvashydromenu.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        manageEnterWindow  
        
        
    elif item_iid == "I016":
        # ===========================Areamenu User  Guide==================================================
        frame1 = Frame(mainframe2)
        f_areamenu = Frame(frame1, bg='blue')
        canvasareamenu = Canvas(f_areamenu, borderwidth=0)
        frameareamenu00 = Frame(canvasareamenu)
        
        def mousewheel(event):canvasareamenu.yview_scroll(int(-1*(event.delta/120)),"units")
        def bind_to_wheel(event):canvasareamenu.bind_all("<MouseWheel>",mousewheel)
        def unbind_to_wheel(event):canvasareamenu.unbind_all("<MouseWheel>")
        f_areamenu.bind('<Enter>',bind_to_wheel)    
        f_areamenu.bind('<Leave>',unbind_to_wheel)  
        
       
        vsb = Scrollbar(f_areamenu, orient="vertical", command=canvasareamenu.yview)
        vsb.pack(side="right", fill="y")
        hsb = Scrollbar(f_areamenu, orient="horizontal", command=canvasareamenu.xview)
        hsb.pack(side="bottom", fill="x")
        #canvasareamenu.pack(side="left", fill="both", expand=True)
        canvasareamenu.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        
        frameareamenu00.bind("<Configure>", lambda event, canvas=canvasareamenu: onFrameConfigure(canvasareamenu))
        canvasareamenu.create_window((1,1), window=frameareamenu00, anchor="nw")
        canvasareamenu.configure(yscrollcommand=vsb.set)
        canvasareamenu.configure(xscrollcommand=hsb.set)
        
        
        imageareamenu1 = PhotoImage(file='Areamenu user guide_001.png')
        #imageareamenu2 = PhotoImage(file='areamenu user guide_002.png')
        imageareamenu3 = PhotoImage(file='Areamenu user guide_003.png')
        
        labelareamenu1 = Label(frameareamenu00, text='''
        Survey Companion "Area" menu has the following submenu:
            
        ⊞ Compute Area: This submenu allows the user to choose a coordinate data entry method to adopt. It 
          has the following functions: 
              
             1. Read coordinates from file: This functionality activates the Area Computation Module in which the
                user is expected to import coordinate files saved in the acceptable file formats
               
             2. Input coordinates directly: This functionality activates the Area Computation Module in which the 
                user is expected to directly input the coordinates of the bounding points.
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        labelareamenuimage1 = Label(frameareamenu00, image=imageareamenu1)
        labelareamenuheading = Label(frameareamenu00, image=imageareamenu3)
        
        labelareamenuheading.pack()
        labelareamenu1.pack()
        labelareamenuimage1.pack()
        frame1.place(x=0,y=0,relheight=1.0,relwidth=1.0)
        f_areamenu.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        canvasareamenu.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        manageEnterWindow  
        
        
    elif item_iid == "I017":
        # ===========================Resectionmenu User  Guide==================================================
        frame1 = Frame(mainframe2)
        f_resectionmenu = Frame(frame1, bg='blue')
        canvasresectionmenu = Canvas(f_resectionmenu, borderwidth=0)
        frameresectionmenu00 = Frame(canvasresectionmenu)
        
        def mousewheel(event):canvasresectionmenu.yview_scroll(int(-1*(event.delta/120)),"units")
        def bind_to_wheel(event):canvasresectionmenu.bind_all("<MouseWheel>",mousewheel)
        def unbind_to_wheel(event):canvasresectionmenu.unbind_all("<MouseWheel>")
        f_resectionmenu.bind('<Enter>',bind_to_wheel)    
        f_resectionmenu.bind('<Leave>',unbind_to_wheel)  
        
      
        vsb = Scrollbar(f_resectionmenu, orient="vertical", command=canvasresectionmenu.yview)
        vsb.pack(side="right", fill="y")
        hsb = Scrollbar(f_resectionmenu, orient="horizontal", command=canvasresectionmenu.xview)
        hsb.pack(side="bottom", fill="x")
        #canvasresectionmenu.pack(side="left", fill="both", expand=True)
        canvasresectionmenu.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        
        frameresectionmenu00.bind("<Configure>", lambda event, canvas=canvasresectionmenu: onFrameConfigure(canvasresectionmenu))
        canvasresectionmenu.create_window((1,1), window=frameresectionmenu00, anchor="nw")
        canvasresectionmenu.configure(yscrollcommand=vsb.set)
        canvasresectionmenu.configure(xscrollcommand=hsb.set)
        
        
        imageresectionmenu1 = PhotoImage(file='Resectionmenu user guide_001.png')
        #imageresectionmenu2 = PhotoImage(file='resectionmenu user guide_002.png')
        imageresectionmenu3 = PhotoImage(file='Resectionmenu user guide_003.png')
        
        labelresectionmenu1 = Label(frameresectionmenu00, text='''
        Survey Companion “Resection” menu activates the Three Point Resection Computation Module
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        labelresectionmenuimage1 = Label(frameresectionmenu00, image=imageresectionmenu1)
        labelresectionmenuheading = Label(frameresectionmenu00, image=imageresectionmenu3)
        
        labelresectionmenuheading.pack()
        labelresectionmenu1.pack()
        labelresectionmenuimage1.pack()
        frame1.place(x=0,y=0,relheight=1.0,relwidth=1.0)
        f_resectionmenu.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        canvasresectionmenu.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        manageEnterWindow  
    
    
    elif item_iid == "I018":
        # ===========================Intersectionmenu User  Guide==================================================
        frame1 = Frame(mainframe2)
        f_intersectionmenu = Frame(frame1)
        canvasintersectionmenu = Canvas(f_intersectionmenu, borderwidth=0)
        frameintersectionmenu00 = Frame(canvasintersectionmenu)
        
        def mousewheel(event):canvasintersectionmenu.yview_scroll(int(-1*(event.delta/120)),"units")
        def bind_to_wheel(event):canvasintersectionmenu.bind_all("<MouseWheel>",mousewheel)
        def unbind_to_wheel(event):canvasintersectionmenu.unbind_all("<MouseWheel>")
        f_intersectionmenu.bind('<Enter>',bind_to_wheel)    
        f_intersectionmenu.bind('<Leave>',unbind_to_wheel)  
        
        
        vsb = Scrollbar(f_intersectionmenu, orient="vertical", command=canvasintersectionmenu.yview)
        vsb.pack(side="right", fill="y")
        hsb = Scrollbar(f_intersectionmenu, orient="horizontal", command=canvasintersectionmenu.xview)
        hsb.pack(side="bottom", fill="x")
        #canvasintersectionmenu.pack(side="left", fill="both", expand=True)
        canvasintersectionmenu.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        
        frameintersectionmenu00.bind("<Configure>", lambda event, canvas=canvasintersectionmenu: onFrameConfigure(canvasintersectionmenu))
        canvasintersectionmenu.create_window((1,1), window=frameintersectionmenu00, anchor="nw")
        canvasintersectionmenu.configure(yscrollcommand=vsb.set)
        canvasintersectionmenu.configure(xscrollcommand=hsb.set)
        
        
        imageintersectionmenu1 = PhotoImage(file='Intersectionmenu user guide_001.png')
        imageintersectionmenu2 = PhotoImage(file='Intersectionmenu user guide_002.png')
        imageintersectionmenu3 = PhotoImage(file='Intersectionmenu user guide_003.png')
        

        labelintersectionmenu1 = Label(frameintersectionmenu00, text='''
        Survey Companion "Intersection" menu has the following submenus:
            
        ⊞ Compute Intersection: This submenu allows the user to choose a type of intersection problem to solve. 
          It has the following functions: 
              
               1. Azimuth_Azimuth Intersection: This function activates the module for azimuth-azimuth (direction-
                  direction) intersection problem.
               
               2. Azimuth_Distance Intersection: This function activates the module for azimuth-distance (direction-
                  distance) intersection problem.
               
               3. Distance_Distance Intersection: This function activates the module for distance-distance 
                  intersection problem.
                  
               4. Angle_Angle Intersection: This function activates the module for angle-angle intersection problem.
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        # =============================================================================
        labelintersectionmenuimage1 = Label(frameintersectionmenu00, image=imageintersectionmenu1)
        
        labelintersectionmenu2 = Label(frameintersectionmenu00, text='''
        ⊞ Compute Side Section: This submenu allows the user to choose a type of side section problem to solve. It has 
          the following functions: 
              
               1. Interior Angle Side section: This function activates the module for interior angle side section 
                  problem.
                  
               2. Distance Side section: This function activates the module for distance side section problem.
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        labelintersectionmenuimage2 = Label(frameintersectionmenu00, image=imageintersectionmenu2)
        labelintersectionmenuheading = Label(frameintersectionmenu00, image=imageintersectionmenu3)
    
        labelintersectionmenuheading.pack()
        labelintersectionmenu1.pack()
        labelintersectionmenuimage1.pack()
        labelintersectionmenu2.pack()
        labelintersectionmenuimage2.pack()
        frame1.place(x=0,y=0,relheight=1.0,relwidth=1.0)
        f_intersectionmenu.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        canvasintersectionmenu.pack(side=LEFT, expand=True,fill=BOTH)
        manageEnterWindow
        
    
    elif item_iid == "I019":
        # ===========================Helpmenu User  Guide==================================================
        frame1 = Frame(mainframe2)
        f_helpmenu = Frame(frame1, bg='blue')
        canvashelpmenu = Canvas(f_helpmenu, borderwidth=0)
        framehelpmenu00 = Frame(canvashelpmenu)
        
        def mousewheel(event):canvashelpmenu.yview_scroll(int(-1*(event.delta/120)),"units")
        def bind_to_wheel(event):canvashelpmenu.bind_all("<MouseWheel>",mousewheel)
        def unbind_to_wheel(event):canvashelpmenu.unbind_all("<MouseWheel>")
        f_helpmenu.bind('<Enter>',bind_to_wheel)    
        f_helpmenu.bind('<Leave>',unbind_to_wheel)  
        
      
        vsb = Scrollbar(f_helpmenu, orient="vertical", command=canvashelpmenu.yview)
        vsb.pack(side="right", fill="y")
        hsb = Scrollbar(f_helpmenu, orient="horizontal", command=canvashelpmenu.xview)
        hsb.pack(side="bottom", fill="x")
        #canvashelpmenu.pack(side="left", fill="both", expand=True)
        canvashelpmenu.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        
        framehelpmenu00.bind("<Configure>", lambda event, canvas=canvashelpmenu: onFrameConfigure(canvashelpmenu))
        canvashelpmenu.create_window((1,1), window=framehelpmenu00, anchor="nw")
        canvashelpmenu.configure(yscrollcommand=vsb.set)
        canvashelpmenu.configure(xscrollcommand=hsb.set)
        
        
        imagehelpmenu1 = PhotoImage(file='Helpmenu user guide_001.png')
        #imagehelpmenu2 = PhotoImage(file='helpmenu user guide_002.png')
        imagehelpmenu3 = PhotoImage(file='Helpmenu user guide_003.png')
        
        labelhelpmenu1 = Label(framehelpmenu00, text='''
        Survey Companion "Help" menu has the following functionalities:
            
        ⊞ About: This functionality opens up the window for Survey Companion information and developer’s address. 
        
        ⊞ User Guide: This functionality opens up the window for Survey Companion user guide. This provides the 
          user with instruction manual for easy use of the software. 
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        labelhelpmenuimage1 = Label(framehelpmenu00, image=imagehelpmenu1)
        labelhelpmenuheading = Label(framehelpmenu00, image=imagehelpmenu3)
        
        labelhelpmenuheading.pack()
        labelhelpmenu1.pack()
        labelhelpmenuimage1.pack()
        frame1.place(x=0,y=0,relheight=1.0,relwidth=1.0)
        f_helpmenu.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        canvashelpmenu.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        manageEnterWindow  
    
    
    elif item_iid == "I006":
        # ===========================Traverse User  Guide==================================================
        frame1 = Frame(mainframe2)
        f_traverse = Frame(frame1)
        canvastraverse = Canvas(f_traverse, borderwidth=0)
        frametraverse00 = Frame(canvastraverse)
        
        def mousewheel(event):canvastraverse.yview_scroll(int(-1*(event.delta/120)),"units")
        def bind_to_wheel(event):canvastraverse.bind_all("<MouseWheel>",mousewheel)
        def unbind_to_wheel(event):canvastraverse.unbind_all("<MouseWheel>")
        f_traverse.bind('<Enter>',bind_to_wheel)    
        f_traverse.bind('<Leave>',unbind_to_wheel)  
        
        
        vsb = Scrollbar(f_traverse, orient="vertical", command=canvastraverse.yview)
        vsb.pack(side="right", fill="y")
        hsb = Scrollbar(f_traverse, orient="horizontal", command=canvastraverse.xview)
        hsb.pack(side="bottom", fill="x")
        #canvastraverse.pack(side="left", fill="both", expand=True)
        canvastraverse.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        
        frametraverse00.bind("<Configure>", lambda event, canvas=canvastraverse: onFrameConfigure(canvastraverse))
        canvastraverse.create_window((1,1), window=frametraverse00, anchor="nw")
        canvastraverse.configure(yscrollcommand=vsb.set)
        canvastraverse.configure(xscrollcommand=hsb.set)
        
        
        imagetraverse1 = PhotoImage(file='Traverse user guide_001.png')
        imagetraverse2 = PhotoImage(file='Traverse user guide_002.png')
        imagetraverse3 = PhotoImage(file='Traverse user guide_003.png')
        
        
        labeltraverse1 = Label(frametraverse00, text='''
        This module can reduce a one-zero traverse field sheet, adjust the traverse network and compute the 
        adjusted coordinates of the traverse stations.
        
        The module handles all types of closed traverse network which includes:
           1. Polygon Traverse (mathematically and geometrically closed network)
           2. Connecting Traverse (mathematically closed but geometrically open network)
           
        The program accepts field data file presented in a Microsoft Excel spreadsheet. The user should ensure 
        that the data file is saved properly using the .xlsx or .xls file extension eg. “Traverse Datafile.xlsx” 
        or “Traverse Datafile.xls”. The user can then open the data file from the file menu in the Menu bar.
        
        For your Excel data files, fill in the Station_At, Face, Horizontal Angles, Distance and Station_To 
        (in this order) in different columns (one for each). Also provide the Starting and Closing Controls 
        and their coordinates in the data file. The controls section of the data file should come with field 
        headers (CONTROLS, STATION ID, EASTINGS, and NORTHINGS). Each field should hold four row values 
        (Start Instrument Station, Start Backsight, Close Instrument Station, and Close Foresight).
        
        While filling the fields in the data file, take note of the following:
            
           1. For every instrument setup, the face left angular observations are to be filled in first before 
              the face right readings. Also, for every face observation, backsight angular readings are observed 
              first before foresight angular reading. That is, for every instrument station, angular observations 
              should come in the order: face left for backsight, face left for foresight, face right for backsight, 
              and face right for foresight.
              
           2. Your Angular observations have to be separated with a space not the conventional degrees, minutes 
              and seconds symbols. For example, input 342°09'58" as 342 09 58
              
           3. Your distance field requires only one value which is the distance from the instrument station to 
              the foresight station
              
           4. Your distance values should be strictly numeric. For example, input 123.456m as 123.456
        
        Below is a typical Traverse Field data that would be accepted by the program.
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        # =============================================================================
        labeltraverseimage1 = Label(frametraverse00, image=imagetraverse1)
        
        labeltraverse2 = Label(frametraverse00, text='''
        Before computing the traverse, the user can check the boxes to either 'Plot the Traverse Network' or 
        'Show Computation Sheet'. Checking either of the boxes provides a visual plot of the network and 
        provides the computation sheet of the network respectively.
        
        Before computing the traverse, the user is also expected to choose a Computation Method: either the 
        Bowditch Rule or the Transit Rule.
        
        Finally, clicking the “COMPUTE” button does the magic.
        
        Upon computation, the results can be exported to different file formats as below:
            
           1. Microsoft Excel Spreadsheet (.xlsx): This export functionality can be done from the “File” menu »  
              “Export output” » ”Export to xlsx”. It exports either the coordinate listing or the computation sheet to 
              an Excel file.
              
           2. Comma Delimited File (.csv): This export functionality can be done from the “File” menu » “Export 
              output” » “Export to csv”. It exports either the coordinate listing or the computation sheet to a comma 
              elimited file which can be opened by Notepad or Microsoft Excel.
              
           3. Shapefile (.shp): This export functionality can be done from the “File” menu” » “Export output” » 
              “Export to shp”. It exports the plot of the traverse network to a shapefile which can be opened by any 
              Geographic Information System (GIS) based software example ArcGIS and QGIS.
              
           4. Drawing Exchange Format (.dxf): This export functionality can be done from the “File” menu » “Export 
              output” » “Export to dxf”. It exports the plot of the traverse network to a drawing exchange file format 
              (.dxf) which can be opened by any Computer Aided Design (CAD) based software example AutoCAD and 
              MicroStation.
              
           5. Raster Image File Formats: This export functionality is only available if the 'Plot the Traverse 
              Network' checkbox was checked before traverse computation, that is, it can only be accessed from the 
              “Plot of the Traverse Network” window. This export can be done from the “Plot of the Traverse Network” 
              toolbar » “Save the figure ( )”. Different image file formats are achievable example png, jpeg, pdf, tif, 
              svg, pgf and so on.
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        labeltraverseimage2 = Label(frametraverse00, image=imagetraverse2)
        labeltraverseheading = Label(frametraverse00, image=imagetraverse3)
    
        labeltraverseheading.pack()
        labeltraverse1.pack()
        labeltraverseimage1.pack()
        labeltraverse2.pack()
        labeltraverseimage2.pack()
        frame1.place(x=0,y=0,relheight=1.0,relwidth=1.0)
        f_traverse.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        canvastraverse.pack(side=LEFT, expand=True,fill=BOTH)
        manageEnterWindow
        
    
    elif item_iid == "I007":
        # ===========================Level User  Guide==================================================
        frame1 = Frame(mainframe2)
        f_level = Frame(frame1, bg='blue')
        canvaslevel = Canvas(f_level, borderwidth=0)
        framelevel00 = Frame(canvaslevel)
        
        def mousewheel(event):canvaslevel.yview_scroll(int(-1*(event.delta/120)),"units")
        def bind_to_wheel(event):canvaslevel.bind_all("<MouseWheel>",mousewheel)
        def unbind_to_wheel(event):canvaslevel.unbind_all("<MouseWheel>")
        f_level.bind('<Enter>',bind_to_wheel)    
        f_level.bind('<Leave>',unbind_to_wheel)  
        
        
        vsb = Scrollbar(f_level, orient="vertical", command=canvaslevel.yview)
        vsb.pack(side="right", fill="y")
        hsb = Scrollbar(f_level, orient="horizontal", command=canvaslevel.xview)
        hsb.pack(side="bottom", fill="x")
        #canvaslevel.pack(side="left", fill="both", expand=True)
        canvaslevel.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        
        framelevel00.bind("<Configure>", lambda event, canvas=canvaslevel: onFrameConfigure(canvaslevel))
        canvaslevel.create_window((1,1), window=framelevel00, anchor="nw")
        canvaslevel.configure(yscrollcommand=vsb.set)
        canvaslevel.configure(xscrollcommand=hsb.set)
        
        
        imagelevel1 = PhotoImage(file='Level user guide_001.png')
        imagelevel2 = PhotoImage(file='Level user guide_002.png')
        imagelevel3 = PhotoImage(file='Level user guide_003.png')
        
        labellevel1 = Label(framelevel00, text='''
        This module can reduce a level network field sheet, adjust the level network (given the heights of 
        the benchmark) and compute the adjusted reduced heights of all level staff stations.
        
        The module handles all types of closed level network which includes:
           1. Loop Level Network 
           2. Connecting Level Network 
           
        The program accepts field data file presented in a Microsoft Excel spreadsheet. The user should ensure 
        that the data file is saved properly using the .xlsx or .xls file extension eg. “Level Datafile.xlsx” 
        or “Level Datafile.xls”. The user can then open the data file from the file menu in the Menu bar.
        
        For your Excel data files, fill in the Staff_Stations, Chainages, Backsights, Intermediate sight and 
        Foresights (in this order) in different columns (one for each). Also provide the Starting and Closing 
        Benchmarks and their actual or assumed heights in the data file. The benchmarks section of the data file
        should come with field headers (BENCHMARKS, STATION ID, and HEIGHT). Each field should hold two row values
        (Start benchmark and Close benchmark).
        
        While filling the fields in the data file, take note of the following:
            
           1. Ensure that the first backsight staff station is your starting benchmark. Also, ensure that the last
              foresight staff station is your closing benchmark.
              
           2. All staff reading values should be strictly numeric. For example, input 1.234m as 1.234
           
        Below is a typical level network field data that would be accepted by the program:
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        # =============================================================================
        labellevelimage1 = Label(framelevel00, image=imagelevel1)
        
        labellevel2 = Label(framelevel00, text='''
        Before reducing the level network, the user can check the box to 'Show Computation Sheet'. Checking the box 
        provides the computation sheet of the network.
        
        Finally, clicking the “REDUCE” button does the magic.
        
        Upon computation, the results can be exported to different file formats as below:
            
           1. Microsoft Excel Spreadsheet (.xlsx): This export functionality can be done from the “File” menu »  
              “Export output” » ”Export to xlsx”. It exports either the reduced heights listing or the computation 
              sheet to an Excel file.
              
           2. Comma Delimited File (.csv): This export functionality can be done from the “File” menu » “Export 
              output” » “Export to csv”. It exports either the reduced heights listing or the computation sheet to 
              a comma delimited file which can be opened by Notepad or Microsoft Excel.
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        labellevelimage2 = Label(framelevel00, image=imagelevel2)
        labellevelheading = Label(framelevel00, image=imagelevel3)
        
        labellevelheading.pack()
        labellevel1.pack()
        labellevelimage1.pack()
        labellevel2.pack()
        labellevelimage2.pack()
        frame1.place(x=0,y=0,relheight=1.0,relwidth=1.0)
        f_level.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        canvaslevel.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        manageEnterWindow
        
        
    elif item_iid == "I008":
        # ===========================Hydro User  Guide==================================================
        frame1 = Frame(mainframe2)
        f_hydro = Frame(frame1, bg='blue')
        canvashydro = Canvas(f_hydro, borderwidth=0)
        framehydro00 = Frame(canvashydro)
        
        def mousewheel(event):canvashydro.yview_scroll(int(-1*(event.delta/120)),"units")
        def bind_to_wheel(event):canvashydro.bind_all("<MouseWheel>",mousewheel)
        def unbind_to_wheel(event):canvashydro.unbind_all("<MouseWheel>")
        f_hydro.bind('<Enter>',bind_to_wheel)    
        f_hydro.bind('<Leave>',unbind_to_wheel)  
        
        
        vsb = Scrollbar(f_hydro, orient="vertical", command=canvashydro.yview)
        vsb.pack(side="right", fill="y")
        hsb = Scrollbar(f_hydro, orient="horizontal", command=canvashydro.xview)
        hsb.pack(side="bottom", fill="x")
        #canvashydro.pack(side="left", fill="both", expand=True)
        canvashydro.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        
        framehydro00.bind("<Configure>", lambda event, canvas=canvashydro: onFrameConfigure(canvashydro))
        canvashydro.create_window((1,1), window=framehydro00, anchor="nw")
        canvashydro.configure(yscrollcommand=vsb.set)
        canvashydro.configure(xscrollcommand=hsb.set)
        
        
        imagehydro1 = PhotoImage(file='Hydro user guide_001.png')
        imagehydro2 = PhotoImage(file='Hydro user guide_002.png')
        imagehydro3 = PhotoImage(file='Hydro user guide_003.png')
        
        labelhydro1 = Label(framehydro00, text='''
        This module can filter off false depths from bathymetry/sounding data, and reduce the data to the 
        chart datum (given that water level was reduced to the chart datum before tide observations).
        
        The module also filters off superfluous fixes from the sounding data using distance criterion:
            
        The program accepts field data file presented in comma delimited (.csv) file formats. The user should
        ensure that the data file is saved properly using the .csv file extension eg. “Hydro Datafile.csv”. 
        The user can then open the data file from the file menu in the Menu bar.
        
        The user will also be required to import tide data file by clicking the “OPEN” button in the “Data Filter”
        pane of the sounding data reduction and filtering module. The tide data file should also be saved in .csv 
        file extension eg. “Tide observation file.csv”.
        
        For the data file, fill in the time values, Eastings, Northings, and Sounded depths (in this order) in 
        comma separated quadrads (one fix for each). For the tide data file, fill in the time values and tide 
        levels in comma separated pairs (one tide observation for each).
        
        While filling the fields in both data file and tide data file, take note of the following:
            
           1. For every time value, the hour, minutes and seconds constituents have to be separated with colon(:),
              that is, in the format, “hh:mm:ss” example, 10:05:47. 
              
           2. All other fields (easting and northing coordinate values, sounded depths and tide levels) are 
              strictly numeric. 
              
        Below is an extract of a typical sounding data and a typical tide observation file that would be accepted 
        by the program:
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        # =============================================================================
        labelhydroimage1 = Label(framehydro00, image=imagehydro1)
        
        labelhydro2 = Label(framehydro00, text='''
        **Note that sounding data filtering is not compulsory. 
        If the user wishes to filter off false depths and superfluous data from the data file, input the 
        corresponding fields in the “Data Filter” pane.
        
        Before reducing the sounding data, the user can check the box to 'Plot the Spot Depths’. Checking the 
        box provides a visual plot of the spot depths.
        
        Finally, clicking the “REDUCE” button does the magic.
        
        Upon reduction, the results can be exported to different file formats as below:
            
           1. Microsoft Excel Spreadsheet (.xlsx): This export functionality can be done from the “File” menu »
              “Export output” » ”Export to xlsx”. It exports the reduced and(or) filtered depths and coordinates 
              of the corresponding fixes to an Excel file.
              
           2. Comma Delimited File (.csv): This export functionality can be done from the “File” menu » “Export 
              output” » “Export to csv”. It exports the reduced and(or) filtered depths and coordinates of the 
              corresponding fixes to an Excel file.
              
           3. Shapefile (.shp): This export functionality can be done from the “File” menu” » “Export output” »
              “Export to shp”. It exports the plot of the spot depths to a shapefile which can be opened by any 
              Geographic Information System (GIS) based software example ArcGIS and QGIS. 
              
           4. Drawing Exchange Format (.dxf): This export functionality can be done from the “File” menu » “Export 
              output” » “Export to dxf”. It exports the plot of the spot depths to a drawing exchange file format 
              (.dxf) which can be opened by any Computer Aided Design (CAD) based software example AutoCAD and 
              MicroStation.
              
           5. Raster Image File Formats: This export functionality is only available if the 'Plot the Spot Depths’
              checkbox was checked before the reduction, that is, it can only be accessed from the “Plot of the 
              Spot Depths” window. This export can be done from the “Plot of the Spot Depths” toolbar » “Save the 
              figure ( )”. Different image file formats are achievable example png, jpeg, pdf, tif, svg, pgf and 
              so on.
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        labelhydroimage2 = Label(framehydro00, image=imagehydro2)
        labelhydroheading = Label(framehydro00, image=imagehydro3)
        
        labelhydroheading.pack()
        labelhydro1.pack()
        labelhydroimage1.pack()
        labelhydro2.pack()
        labelhydroimage2.pack()
        frame1.place(x=0,y=0,relheight=1.0,relwidth=1.0)
        f_hydro.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        canvashydro.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        manageEnterWindow


    elif item_iid == "I009":
        # ===========================Areahome User  Guide==================================================
        frame1 = Frame(mainframe2)
        f_areahome = Frame(frame1, bg='blue')
        canvasareahome = Canvas(f_areahome, borderwidth=0)
        frameareahome00 = Frame(canvasareahome)
        
        def mousewheel(event):canvasareahome.yview_scroll(int(-1*(event.delta/120)),"units")
        def bind_to_wheel(event):canvasareahome.bind_all("<MouseWheel>",mousewheel)
        def unbind_to_wheel(event):canvasareahome.unbind_all("<MouseWheel>")
        f_areahome.bind('<Enter>',bind_to_wheel)    
        f_areahome.bind('<Leave>',unbind_to_wheel)  
        
        
        vsb = Scrollbar(f_areahome, orient="vertical", command=canvasareahome.yview)
        vsb.pack(side="right", fill="y")
        hsb = Scrollbar(f_areahome, orient="horizontal", command=canvasareahome.xview)
        hsb.pack(side="bottom", fill="x")
        #canvasareahome.pack(side="left", fill="both", expand=True)
        canvasareahome.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        
        frameareahome00.bind("<Configure>", lambda event, canvas=canvasareahome: onFrameConfigure(canvasareahome))
        canvasareahome.create_window((1,1), window=frameareahome00, anchor="nw")
        canvasareahome.configure(yscrollcommand=vsb.set)
        canvasareahome.configure(xscrollcommand=hsb.set)
        
        
        #imageareahome1 = PhotoImage(file='Areahome user guide_001.png')
        #imageareahome2 = PhotoImage(file='Areahome user guide_002.png')
        imageareahome3 = PhotoImage(file='Areahome user guide_003.png')
        
        labelareahome1 = Label(frameareahome00, text='''
        This module can compute the area and the perimeter of a closed region given the coordinates of the 
        bounding points. 

        The program accepts either reading coordinates from a data file or direct input of the bounding point
        coordinates
        
        While filling the coordinate fields to be used by this module, the user should note the following:
            
           1. The coordinates should be filled in consecutively either in a clockwise or anti-clockwise 
              order.
           2. All fields (easting and northing coordinate values) are strictly numeric. 
           
        **Note that repeating the first coordinate to close the region is not compulsory. 
        
        Before computing the area, the user can check the box to 'Plot the Enclosed Region'. Checking the box 
        provides a visual plot of the enclosed region.
        
        Finally, clicking the “COMPUTE” button does the magic.
        
        Upon computation, the results can be exported to different file formats as below:
            
           1. Shapefile (.shp): This export functionality can be done from the “File” menu” » “Export output” »
              “Export to shp”. It exports the plot of the enclosed region to a shapefile which can be opened by
              any Geographic Information System (GIS) based software example ArcGIS and QGIS.
              
           2. Drawing Exchange Format (.dxf): This export functionality can be done from the “File” menu » “Export 
              output” » “Export to dxf”. It exports the plot of the enclosed region to a drawing exchange file 
              format (.dxf) which can be opened by any Computer Aided Design (CAD) based software example AutoCAD
              and MicroStation.
              
           3. Text file (.txt): This export functionality can be done from the “File” menu »  “Export output” »
              ”Export to txt”. It exports the generated area and perimeter report to text file format (.txt) 
              which can be opened by any text-editing software example Notepad.
        
           4. Raster Image File Formats: This export functionality is only available if the 'Plot the Enclosed 
              Region' checkbox was checked before the reduction, that is, it can only be accessed from the “Plot 
              of the Spot Depths” window. This export can be done from the “Plot of the Spot Depths” toolbar » 
              “Save the figure ( )”. Different image file formats are achievable example png, jpeg, pdf, tif, svg,
              pgf and so on.
        ''',justify='left',font=('helvetica', 10,'italic'))
        

        labelareahomeheading = Label(frameareahome00, image=imageareahome3)
        
        labelareahomeheading.pack()
        labelareahome1.pack()
        frame1.place(x=0,y=0,relheight=1.0,relwidth=1.0)
        f_areahome.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        canvasareahome.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        manageEnterWindow


    elif item_iid == "I00A":
        # ===========================Areareadfile User  Guide==================================================
        frame1 = Frame(mainframe2)
        f_areareadfile = Frame(frame1, bg='blue')
        canvasareareadfile = Canvas(f_areareadfile, borderwidth=0)
        frameareareadfile00 = Frame(canvasareareadfile)
        
        def mousewheel(event):canvasareareadfile.yview_scroll(int(-1*(event.delta/120)),"units")
        def bind_to_wheel(event):canvasareareadfile.bind_all("<MouseWheel>",mousewheel)
        def unbind_to_wheel(event):canvasareareadfile.unbind_all("<MouseWheel>")
        f_areareadfile.bind('<Enter>',bind_to_wheel)    
        f_areareadfile.bind('<Leave>',unbind_to_wheel)  
        
        
        vsb = Scrollbar(f_areareadfile, orient="vertical", command=canvasareareadfile.yview)
        vsb.pack(side="right", fill="y")
        hsb = Scrollbar(f_areareadfile, orient="horizontal", command=canvasareareadfile.xview)
        hsb.pack(side="bottom", fill="x")
        #canvasareareadfile.pack(side="left", fill="both", expand=True)
        canvasareareadfile.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        
        frameareareadfile00.bind("<Configure>", lambda event, canvas=canvasareareadfile: onFrameConfigure(canvasareareadfile))
        canvasareareadfile.create_window((1,1), window=frameareareadfile00, anchor="nw")
        canvasareareadfile.configure(yscrollcommand=vsb.set)
        canvasareareadfile.configure(xscrollcommand=hsb.set)
        
        
        imageareareadfile1 = PhotoImage(file='Areareadfile user guide_001.png')
        imageareareadfile2 = PhotoImage(file='Areareadfile user guide_002.png')
        imageareareadfile3 = PhotoImage(file='Areareadfile user guide_003.png')
        
        labelareareadfile1 = Label(frameareareadfile00, text='''
        This sub-module in the area computation module accepts import of coordinate file for computation of 
        the area and the perimeter of a closed region given the coordinates of the bounding points. 

        The coordinate file that is accepted by this sub-module includes files in the following formats:
            
           1. Microsoft Excel files (.xlsx): The user should ensure that the coordinate file is saved properly
              using the .xlsx or .xls file extension eg. “Area Coordinate file.xlsx” or “Area Coordinate 
              file.xls”. The user can then open the coordinate file from the file menu in the Menu bar.
              
              For your Excel coordinate files, fill in the easting and northing coordinate values (in this 
              order) in different columns (one for each).
              
              ** Note that the columns should come with column headers.
              
           2. Comma delimited file (.csv): The user should ensure that the coordinate file is saved properly
              using the .csv file extension eg. “Area Coordinate file.csv”. The user can then open the coordinate
              file from the file menu in the Menu bar.
              
              For the comma delimited file, fill in the easting and northing coordinate values in comma separated
              pairs (new line for each).
              
           3. Text file (.txt): The user should ensure that the coordinate file is saved properly using the .txt 
              file extension eg. “Area Coordinate file.txt”. The user can then open the coordinate file from the 
              file menu in the Menu bar.
              
        For the text file, fill in the easting and northing coordinate values in comma separated pairs (new line 
        for each).
        
        Below are typical coordinate files in the different file formats that would be accepted by the program:
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        # =============================================================================
        labelareareadfileimage1 = Label(frameareareadfile00, image=imageareareadfile1)
        
        labelareareadfile2 = Label(frameareareadfile00, text='''
        While filling the coordinate fields to be used by this module, the user should note the following:
            
           1. The coordinates should be filled in consecutively either in a clockwise or anti-clockwise 
              order.
              
           2. All fields (easting and northing coordinate values) are strictly numeric. 
           
        **Note that repeating the first coordinate to close the region is not compulsory. 
        
        Before computing the area, the user can check the box to 'Plot the Enclosed Region'. Checking the box 
        provides a visual plot of the enclosed region.
        
        Finally, clicking the “COMPUTE” button does the magic.
        
        Upon computation, the results can be exported to different file formats as below:
            
           1. Shapefile (.shp): This export functionality can be done from the “File” menu” » “Export output” »
              “Export to shp”. It exports the plot of the enclosed region to a shapefile which can be opened by
              any Geographic Information System (GIS) based software example ArcGIS and QGIS.
              
           2. Drawing Exchange Format (.dxf): This export functionality can be done from the “File” menu » “Export 
              output” » “Export to dxf”. It exports the plot of the enclosed region to a drawing exchange file 
              format (.dxf) which can be opened by any Computer Aided Design (CAD) based software example AutoCAD
              and MicroStation.
              
           3. Text file (.txt): This export functionality can be done from the “File” menu »  “Export output” »
              ”Export to txt”. It exports the generated area and perimeter report to text file format (.txt) 
              which can be opened by any text-editing software example Notepad.
        
           4. Raster Image File Formats: This export functionality is only available if the 'Plot the Enclosed 
              Region' checkbox was checked before the reduction, that is, it can only be accessed from the “Plot 
              of the Spot Depths” window. This export can be done from the “Plot of the Spot Depths” toolbar » 
              “Save the figure ( )”. Different image file formats are achievable example png, jpeg, pdf, tif, svg,
              pgf and so on.
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        labelareareadfileimage2 = Label(frameareareadfile00, image=imageareareadfile2)
        labelareareadfileheading = Label(frameareareadfile00, image=imageareareadfile3)
        
        labelareareadfileheading.pack()
        labelareareadfile1.pack()
        labelareareadfileimage1.pack()
        labelareareadfile2.pack()
        labelareareadfileimage2.pack()
        frame1.place(x=0,y=0,relheight=1.0,relwidth=1.0)
        f_areareadfile.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        canvasareareadfile.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        manageEnterWindow
        
        
    elif item_iid == "I00B":
        # ===========================Areadirectinput User  Guide==================================================
        frame1 = Frame(mainframe2)
        f_areadirectinput = Frame(frame1, bg='blue')
        canvasareadirectinput = Canvas(f_areadirectinput, borderwidth=0)
        frameareadirectinput00 = Frame(canvasareadirectinput)
        
        def mousewheel(event):canvasareadirectinput.yview_scroll(int(-1*(event.delta/120)),"units")
        def bind_to_wheel(event):canvasareadirectinput.bind_all("<MouseWheel>",mousewheel)
        def unbind_to_wheel(event):canvasareadirectinput.unbind_all("<MouseWheel>")
        f_areadirectinput.bind('<Enter>',bind_to_wheel)    
        f_areadirectinput.bind('<Leave>',unbind_to_wheel)  
        
        
        vsb = Scrollbar(f_areadirectinput, orient="vertical", command=canvasareadirectinput.yview)
        vsb.pack(side="right", fill="y")
        hsb = Scrollbar(f_areadirectinput, orient="horizontal", command=canvasareadirectinput.xview)
        hsb.pack(side="bottom", fill="x")
        #canvasareadirectinput.pack(side="left", fill="both", expand=True)
        canvasareadirectinput.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        
        frameareadirectinput00.bind("<Configure>", lambda event, canvas=canvasareadirectinput: onFrameConfigure(canvasareadirectinput))
        canvasareadirectinput.create_window((1,1), window=frameareadirectinput00, anchor="nw")
        canvasareadirectinput.configure(yscrollcommand=vsb.set)
        canvasareadirectinput.configure(xscrollcommand=hsb.set)
        
        
        imageareadirectinput1 = PhotoImage(file='Areadirectinput user guide_001.png')
        imageareadirectinput2 = PhotoImage(file='Areadirectinput user guide_002.png')
        imageareadirectinput3 = PhotoImage(file='Areadirectinput user guide_003.png')
        
        labelareadirectinput1 = Label(frameareadirectinput00, text='''
        This sub-module in the area computation module accepts direct coordinate input for computation of 
        the area and the perimeter of a closed region given the coordinates of the bounding points. 

        The coordinates should be inputted in the “Data Input” pane of the program.
        
        For direct coordinate input, fill in the easting and northing coordinate values in comma separated
        pairs (new line for each).
        
        Below is a typical coordinate input that would be accepted by the program:
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        # =============================================================================
        labelareadirectinputimage1 = Label(frameareadirectinput00, image=imageareadirectinput1)
        
        labelareadirectinput2 = Label(frameareadirectinput00, text='''
        While filling the coordinate fields to be used by this module, the user should note the following:
            
           1. The coordinates should be filled in consecutively either in a clockwise or anti-clockwise 
              order.
              
           2. All fields (easting and northing coordinate values) are strictly numeric. 
           
        **Note that repeating the first coordinate to close the region is not compulsory. 
        
        Before computing the area, the user can check the box to 'Plot the Enclosed Region'. Checking the box 
        provides a visual plot of the enclosed region.
        
        Finally, clicking the “COMPUTE” button does the magic.
        
        Upon computation, the results can be exported to different file formats as below:
            
           1. Shapefile (.shp): This export functionality can be done from the “File” menu” » “Export output” »
              “Export to shp”. It exports the plot of the enclosed region to a shapefile which can be opened by
              any Geographic Information System (GIS) based software example ArcGIS and QGIS.
              
           2. Drawing Exchange Format (.dxf): This export functionality can be done from the “File” menu » “Export 
              output” » “Export to dxf”. It exports the plot of the enclosed region to a drawing exchange file 
              format (.dxf) which can be opened by any Computer Aided Design (CAD) based software example AutoCAD
              and MicroStation.
              
           3. Text file (.txt): This export functionality can be done from the “File” menu »  “Export output” »
              ”Export to txt”. It exports the generated area and perimeter report to text file format (.txt) 
              which can be opened by any text-editing software example Notepad.
        
           4. Raster Image File Formats: This export functionality is only available if the 'Plot the Enclosed 
              Region' checkbox was checked before the reduction, that is, it can only be accessed from the “Plot 
              of the Spot Depths” window. This export can be done from the “Plot of the Spot Depths” toolbar » 
              “Save the figure ( )”. Different image file formats are achievable example png, jpeg, pdf, tif, svg,
              pgf and so on.
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        labelareadirectinputimage2 = Label(frameareadirectinput00, image=imageareadirectinput2)
        labelareadirectinputheading = Label(frameareadirectinput00, image=imageareadirectinput3)
        
        labelareadirectinputheading.pack()
        labelareadirectinput1.pack()
        labelareadirectinputimage1.pack()
        labelareadirectinput2.pack()
        labelareadirectinputimage2.pack()
        frame1.place(x=0,y=0,relheight=1.0,relwidth=1.0)
        f_areadirectinput.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        canvasareadirectinput.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        manageEnterWindow       
        
        
    elif item_iid == "I00C":
        # ===========================Resection User  Guide==================================================
        frame1 = Frame(mainframe2)
        f_resection = Frame(frame1, bg='blue')
        canvasresection = Canvas(f_resection, borderwidth=0)
        frameresection00 = Frame(canvasresection)
        
        def mousewheel(event):canvasresection.yview_scroll(int(-1*(event.delta/120)),"units")
        def bind_to_wheel(event):canvasresection.bind_all("<MouseWheel>",mousewheel)
        def unbind_to_wheel(event):canvasresection.unbind_all("<MouseWheel>")
        f_resection.bind('<Enter>',bind_to_wheel)    
        f_resection.bind('<Leave>',unbind_to_wheel)  
        
        
        vsb = Scrollbar(f_resection, orient="vertical", command=canvasresection.yview)
        vsb.pack(side="right", fill="y")
        hsb = Scrollbar(f_resection, orient="horizontal", command=canvasresection.xview)
        hsb.pack(side="bottom", fill="x")
        #canvasresection.pack(side="left", fill="both", expand=True)
        canvasresection.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        
        frameresection00.bind("<Configure>", lambda event, canvas=canvasresection: onFrameConfigure(canvasresection))
        canvasresection.create_window((1,1), window=frameresection00, anchor="nw")
        canvasresection.configure(yscrollcommand=vsb.set)
        canvasresection.configure(xscrollcommand=hsb.set)
        
        
        imageresection1 = PhotoImage(file='Resection user guide_001.png')
        imageresection2 = PhotoImage(file='Resection user guide_002.png')
        imageresection3 = PhotoImage(file='Resection user guide_003.png')
        
        labelresection1 = Label(frameresection00, text='''
        This module can solve the three point resection problem of position fixing, given the coordinates of 
        three control stations and the interior angles observed at the unknown station. 
        
        The program accepts direct input of all data necessary for the computation.
        
        The labels of the control stations are assumed “Control_A”, “Control_B”, and “Control_C” while the 
        unknown station label is assumed “Point_O”
        
        The data required from the user includes the following:
            
           1. Control_A: The coordinate value of the first control point.
           
           2. Control_B: The coordinate value of the second control point.
           
           3. Control_C: The coordinate value of the third control point.
           
           4. Angle AÔB: The clockwise angle between the directions OA and OB
           
           5. Angle BÔC: The clockwise angle between the directions OB and OC
           
        While filling the data fields to be used by this module, the user should note the following:
            
           1. All coordinate value fields are numeric, entered in easting and northing pairs, and separated by 
              a comma, For example,  123456.789,987654.321.
              
           2. All angular fields are numeric and in the form “dd mm ss”, with the degree, minutes and seconds 
              values separated by space. For example, input 342°09'58" as 342 09 58
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        # =============================================================================
        labelresectionimage1 = Label(frameresection00, image=imageresection1)
        
        labelresection2 = Label(frameresection00, text='''
        Before computing the resection, the user can check the box to 'Plot the Resected Point'. Checking the box 
        provides a visual plot of the resected point relative to the control stations.
        
        Finally, clicking the “COMPUTE” button does the magic.
        
        Upon computation, the results can be exported to different file formats as below:
            
           1. Shapefile (.shp): This export functionality can be done from the “File” menu” » “Export output” » 
              “Export to shp”. It exports the plot of the resected point relative to the control stations to a 
              shapefile which can be opened by any Geographic Information System (GIS) based software example ArcGIS
              and QGIS. 
              
           2. Drawing Exchange Format (.dxf): This export functionality can be done from the “File” menu » “Export 
              output” » “Export to dxf”. It exports the plot of the resected point relative to the control stations 
              to a drawing exchange file format (.dxf) which can be opened by any Computer Aided Design (CAD) based 
              software example AutoCAD and IntelliCAD. 
              
           3. Text file (.txt): This export functionality can be done from the “File” menu »  “Export output” » ”Export
              to txt”. It exports the generated coordinate of the unknown point and the additional information in the 
              report to text file format (.txt) which can be opened by any text-editing software example Notepad.
        
           4. Raster Image File Formats: This export functionality is only available if the “Plot the Resected Point” 
              checkbox was checked before the reduction, that is, it can only be accessed from the “Plot the Resected 
              Point” window. This export can be done from the “Plot the Resected Point” toolbar » “Save the figure ( )”. 
              Different image file formats are achievable example png, jpeg, pdf, tif, svg, pgf and so on.
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        labelresectionimage2 = Label(frameresection00, image=imageresection2)
        labelresectionheading = Label(frameresection00, image=imageresection3)
        
        labelresectionheading.pack()
        labelresection1.pack()
        labelresectionimage1.pack()
        labelresection2.pack()
        labelresectionimage2.pack()
        frame1.place(x=0,y=0,relheight=1.0,relwidth=1.0)
        f_resection.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        canvasresection.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        manageEnterWindow  


    elif item_iid == "I00D":
        # ===========================Intersectionhome User  Guide==================================================
        frame1 = Frame(mainframe2)
        f_intersectionhome = Frame(frame1, bg='blue')
        canvasintersectionhome = Canvas(f_intersectionhome, borderwidth=0)
        frameintersectionhome00 = Frame(canvasintersectionhome)
        
        def mousewheel(event):canvasintersectionhome.yview_scroll(int(-1*(event.delta/120)),"units")
        def bind_to_wheel(event):canvasintersectionhome.bind_all("<MouseWheel>",mousewheel)
        def unbind_to_wheel(event):canvasintersectionhome.unbind_all("<MouseWheel>")
        f_intersectionhome.bind('<Enter>',bind_to_wheel)    
        f_intersectionhome.bind('<Leave>',unbind_to_wheel)  
        
        
        vsb = Scrollbar(f_intersectionhome, orient="vertical", command=canvasintersectionhome.yview)
        vsb.pack(side="right", fill="y")
        hsb = Scrollbar(f_intersectionhome, orient="horizontal", command=canvasintersectionhome.xview)
        hsb.pack(side="bottom", fill="x")
        #canvasintersectionhome.pack(side="left", fill="both", expand=True)
        canvasintersectionhome.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        
        frameintersectionhome00.bind("<Configure>", lambda event, canvas=canvasintersectionhome: onFrameConfigure(canvasintersectionhome))
        canvasintersectionhome.create_window((1,1), window=frameintersectionhome00, anchor="nw")
        canvasintersectionhome.configure(yscrollcommand=vsb.set)
        canvasintersectionhome.configure(xscrollcommand=hsb.set)
        
        
        #imageintersectionhome1 = PhotoImage(file='Intersectionhome user guide_001.png')
        #imageintersectionhome2 = PhotoImage(file='Intersectionhome user guide_002.png')
        imageintersectionhome3 = PhotoImage(file='Intersectionhome user guide_003.png')
        
        labelintersectionhome1 = Label(frameintersectionhome00, text='''
        This module can solve the intersection problem of position fixing, given the coordinates of two control 
        stations and any combination of the distance, angle or azimuth of the figure formed by the controls and 
        the unknown point.
        
        The program accepts direct input of all data necessary for the computation.
        
        The labels of the control stations are assumed “Control_A” and “Control_B” while the unknown station 
        label is assumed “Point_O”
        
        Six different types of Intersection problem are possible in this module and they include the following:
            
           1. Azimuth-Azimuth Intersection
           
           2. Azimuth_Distance Intersection
           
           3. Distance_Distance Intersection
           
           4. Angle_Angle Intersection
           
           5. Interior Angle Side section
           
           6. Distance Side section
        
        While filling the data fields to be used by this module, the user should note the following:
            
           1. All coordinate value fields are numeric, entered in easting and northing pairs, and separated by a
              comma, For example,  123456.789,987654.321.
              
           2. All angular fields are numeric and in the form “dd mm ss”, with the degree, minutes and seconds  
              values separated by space. For example, input 342°09'58" as 342 09 58
              
        Before computing the intersection, the user can check the box to 'Plot the Intersected Point'. Checking 
        the box provides a visual plot of the intersected point relative to the control stations.
        
        Finally, clicking the “COMPUTE” button does the magic.
        
        Upon computation, the results can be exported to different file formats as below:
            
        1. Shapefile (.shp): This export functionality can be done from the “File” menu” » “Export output” » 
           “Export to shp”. It exports the plot of the intersected point relative to the control stations to a 
           shapefile which can be opened by any Geographic Information System (GIS) based software example ArcGIS 
           and QGIS.
           
        2. Drawing Exchange Format (.dxf): This export functionality can be done from the “File” menu » “Export output” »
           “Export to dxf”. It exports the plot of the intersected point relative to the control stations to a drawing 
           exchange file format (.dxf) which can be opened by any Computer Aided Design (CAD) based software example 
           AutoCAD and IntelliCAD.
           
        3. Text file (.txt): This export functionality can be done from the “File” menu »  “Export output” » ”Export 
           to txt”. It exports the generated coordinate of the unknown point and the additional information in the 
           report to text file format (.txt) which can be opened by any text-editing software example Notepad.
        
        4. Raster Image File Formats: This export functionality is only available if the “Plot the Intersected Point” 
           checkbox was checked before the reduction, that is, it can only be accessed from the “Plot the Intersected 
           Point” window. This export can be done from the “Plot the Intersected Point” toolbar » “Save the figure ( )”.
           Different image file formats are achievable example png, jpeg, pdf, tif, svg, pgf and so on.
        ''',justify='left',font=('helvetica', 10,'italic'))
        

        labelintersectionhomeheading = Label(frameintersectionhome00, image=imageintersectionhome3)
        
        labelintersectionhomeheading.pack()
        labelintersectionhome1.pack()
        frame1.place(x=0,y=0,relheight=1.0,relwidth=1.0)
        f_intersectionhome.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        canvasintersectionhome.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        manageEnterWindow            


    elif item_iid == "I00E":
        # ===========================Intersectionaz_az User  Guide==================================================
        frame1 = Frame(mainframe2)
        f_intersectionaz_az = Frame(frame1, bg='blue')
        canvasintersectionaz_az = Canvas(f_intersectionaz_az, borderwidth=0)
        frameintersectionaz_az00 = Frame(canvasintersectionaz_az)
        
        def mousewheel(event):canvasintersectionaz_az.yview_scroll(int(-1*(event.delta/120)),"units")
        def bind_to_wheel(event):canvasintersectionaz_az.bind_all("<MouseWheel>",mousewheel)
        def unbind_to_wheel(event):canvasintersectionaz_az.unbind_all("<MouseWheel>")
        f_intersectionaz_az.bind('<Enter>',bind_to_wheel)    
        f_intersectionaz_az.bind('<Leave>',unbind_to_wheel)  
        
        
        vsb = Scrollbar(f_intersectionaz_az, orient="vertical", command=canvasintersectionaz_az.yview)
        vsb.pack(side="right", fill="y")
        hsb = Scrollbar(f_intersectionaz_az, orient="horizontal", command=canvasintersectionaz_az.xview)
        hsb.pack(side="bottom", fill="x")
        #canvasintersectionaz_az.pack(side="left", fill="both", expand=True)
        canvasintersectionaz_az.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        
        frameintersectionaz_az00.bind("<Configure>", lambda event, canvas=canvasintersectionaz_az: onFrameConfigure(canvasintersectionaz_az))
        canvasintersectionaz_az.create_window((1,1), window=frameintersectionaz_az00, anchor="nw")
        canvasintersectionaz_az.configure(yscrollcommand=vsb.set)
        canvasintersectionaz_az.configure(xscrollcommand=hsb.set)
        
        
        imageintersectionaz_az1 = PhotoImage(file='Intersectionaz_az user guide_001.png')
        imageintersectionaz_az2 = PhotoImage(file='Intersectionaz_az user guide_002.png')
        imageintersectionaz_az3 = PhotoImage(file='Intersectionaz_az user guide_003.png')
        
        labelintersectionaz_az1 = Label(frameintersectionaz_az00, text='''
        This sub-module can solve the azimuth-azimuth (direction-direction) intersection problem of position 
        fixing, given the coordinates of two control stations and the azimuths of the intersecting lines
 
        The program accepts direct input of all data necessary for the computation.
        
        The labels of the control stations are assumed “Control_A” and “Control_B” while the unknown station 
        label is assumed “Point_O”
        
        The data required from the user includes the following:
            
           1. Control_A: The coordinate value of the first control point.
           
           2. Control_B: The coordinate value of the second control point.
           
           3. Azimuth AO: The azimuth or whole-circle bearing of the line from the first control point to the   
              unknown point (az_AO)
              
           4. Azimuth BO: The azimuth or whole-circle bearing of the line from the second control point to the 
              unknown point (az_BO)
        
        While filling the data fields to be used by this sub-module, the user should note the following:
            
           1. All coordinate value fields are numeric, entered in easting and northing pairs, and separated by a
              comma, For example,  123456.789,987654.321.
              
           2. All azimuth fields are numeric and in the form “dd mm ss”, with the degree, minutes and seconds  
              values separated by space. For example, input 342°09'58" as 342 09 58
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        # =============================================================================
        labelintersectionaz_azimage1 = Label(frameintersectionaz_az00, image=imageintersectionaz_az1)
        
        labelintersectionaz_az2 = Label(frameintersectionaz_az00, text='''
        Before computing the intersection, the user can check the box to 'Plot the Intersected Point'. Checking the 
        box provides a visual plot of the intersected point relative to the control stations.
        
        Finally, clicking the “COMPUTE” button does the magic.
        
        Upon computation, the results can be exported to different file formats as below:
            
           1. Shapefile (.shp): This export functionality can be done from the “File” menu” » “Export output” » “Export
              to shp”. It exports the plot of the intersected point relative to the control stations to a shapefile 
              which can be opened by any Geographic Information System (GIS) based software example ArcGIS and QGIS.
              
           2. Drawing Exchange Format (.dxf): This export functionality can be done from the “File” menu » “Export output” »
              “Export to dxf”. It exports the plot of the intersected point relative to the control stations to a drawing 
              exchange file format (.dxf) which can be opened by any Computer Aided Design (CAD) based software example 
              AutoCAD and IntelliCAD.
              
           3. Text file (.txt): This export functionality can be done from the “File” menu »  “Export output” » ”Export to
              txt”. It exports the generated coordinate of the unknown point and the additional information in the report 
              to text file format (.txt) which can be opened by any text-editing software example Notepad.
        
           4. Raster Image File Formats: This export functionality is only available if the “Plot the Intersected Point” 
              checkbox was checked before the reduction, that is, it can only be accessed from the “Plot the Intersected 
              Point” window. This export can be done from the “Plot the Intersected Point” toolbar » “Save the figure ( )”.
              Different image file formats are achievable example png, jpeg, pdf, tif, svg, pgf and so on.
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        labelintersectionaz_azimage2 = Label(frameintersectionaz_az00, image=imageintersectionaz_az2)
        labelintersectionaz_azheading = Label(frameintersectionaz_az00, image=imageintersectionaz_az3)
        
        labelintersectionaz_azheading.pack()
        labelintersectionaz_az1.pack()
        labelintersectionaz_azimage1.pack()
        labelintersectionaz_az2.pack()
        labelintersectionaz_azimage2.pack()
        frame1.place(x=0,y=0,relheight=1.0,relwidth=1.0)
        f_intersectionaz_az.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        canvasintersectionaz_az.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        manageEnterWindow  
        
        
    elif item_iid == "I00F":
        # ===========================Intersectionaz_di User  Guide==================================================
        frame1 = Frame(mainframe2)
        f_intersectionaz_di = Frame(frame1, bg='blue')
        canvasintersectionaz_di = Canvas(f_intersectionaz_di, borderwidth=0)
        frameintersectionaz_di00 = Frame(canvasintersectionaz_di)
        
        def mousewheel(event):canvasintersectionaz_di.yview_scroll(int(-1*(event.delta/120)),"units")
        def bind_to_wheel(event):canvasintersectionaz_di.bind_all("<MouseWheel>",mousewheel)
        def unbind_to_wheel(event):canvasintersectionaz_di.unbind_all("<MouseWheel>")
        f_intersectionaz_di.bind('<Enter>',bind_to_wheel)    
        f_intersectionaz_di.bind('<Leave>',unbind_to_wheel)  
        
        
        vsb = Scrollbar(f_intersectionaz_di, orient="vertical", command=canvasintersectionaz_di.yview)
        vsb.pack(side="right", fill="y")
        hsb = Scrollbar(f_intersectionaz_di, orient="horizontal", command=canvasintersectionaz_di.xview)
        hsb.pack(side="bottom", fill="x")
        #canvasintersectionaz_di.pack(side="left", fill="both", expand=True)
        canvasintersectionaz_di.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        
        frameintersectionaz_di00.bind("<Configure>", lambda event, canvas=canvasintersectionaz_di: onFrameConfigure(canvasintersectionaz_di))
        canvasintersectionaz_di.create_window((1,1), window=frameintersectionaz_di00, anchor="nw")
        canvasintersectionaz_di.configure(yscrollcommand=vsb.set)
        canvasintersectionaz_di.configure(xscrollcommand=hsb.set)
        
        
        imageintersectionaz_di1 = PhotoImage(file='Intersectionaz_di user guide_001.png')
        imageintersectionaz_di2 = PhotoImage(file='Intersectionaz_di user guide_002.png')
        imageintersectionaz_di3 = PhotoImage(file='Intersectionaz_di user guide_003.png')
        
        labelintersectionaz_di1 = Label(frameintersectionaz_di00, text='''
        This sub-module can solve the azimuth-distance (direction-distance) intersection problem of position 
        fixing, given the coordinates of two control stations and the azimuth of one intersecting line and the
        distance of the other.
 
        The program accepts direct input of all data necessary for the computation.
        
        The labels of the control stations are assumed “Control_A” and “Control_B” while the unknown station 
        label is assumed “Point_O”
        
        **Note that the control station for which the azimuth to Point_O is available is considered Control_A 
        (first control)
        
        **Note that the control station for which the distance to Point_O is available is considered Control_B 
        (second control)
        
        The data required from the user includes the following:
            
           1. Control_A: The coordinate value of the first control point.
           
           2. Control_B: The coordinate value of the second control point.
           
           3. Azimuth AO: The azimuth or whole-circle bearing of the line from the first control point to the
              unknown point
              
           4. Distance BO: The distance of the line from the second control point to the unknown point

        While filling the data fields to be used by this sub-module, the user should note the following:
            
           1. All coordinate value fields are numeric, entered in easting and northing pairs, and separated by a
              comma, For example,  123456.789,987654.321.
              
           2. All azimuth fields are numeric and in the form “dd mm ss”, with the degree, minutes and seconds  
              values separated by space. For example, input 342°09'58" as 342 09 58
              
           3. All distance fields are numeric. For example, input 123.456m as 123.456
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        # =============================================================================
        labelintersectionaz_diimage1 = Label(frameintersectionaz_di00, image=imageintersectionaz_di1)
        
        labelintersectionaz_di2 = Label(frameintersectionaz_di00, text='''
        Before computing the intersection, the user can check the box to 'Plot the Intersected Point'. Checking the 
        box provides a visual plot of the intersected point relative to the control stations.
        
        Finally, clicking the “COMPUTE” button does the magic.
        
        Upon computation, the results can be exported to different file formats as below:
            
           1. Shapefile (.shp): This export functionality can be done from the “File” menu” » “Export output” » “Export
              to shp”. It exports the plot of the intersected point relative to the control stations to a shapefile 
              which can be opened by any Geographic Information System (GIS) based software example ArcGIS and QGIS.
              
           2. Drawing Exchange Format (.dxf): This export functionality can be done from the “File” menu » “Export output” »
              “Export to dxf”. It exports the plot of the intersected point relative to the control stations to a drawing 
              exchange file format (.dxf) which can be opened by any Computer Aided Design (CAD) based software example 
              AutoCAD and IntelliCAD.
              
           3. Text file (.txt): This export functionality can be done from the “File” menu »  “Export output” » ”Export to
              txt”. It exports the generated coordinate of the unknown point and the additional information in the report 
              to text file format (.txt) which can be opened by any text-editing software example Notepad.
        
           4. Raster Image File Formats: This export functionality is only available if the “Plot the Intersected Point” 
              checkbox was checked before the reduction, that is, it can only be accessed from the “Plot the Intersected 
              Point” window. This export can be done from the “Plot the Intersected Point” toolbar » “Save the figure ( )”.
              Different image file formats are achievable example png, jpeg, pdf, tif, svg, pgf and so on.
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        labelintersectionaz_diimage2 = Label(frameintersectionaz_di00, image=imageintersectionaz_di2)
        labelintersectionaz_diheading = Label(frameintersectionaz_di00, image=imageintersectionaz_di3)
        
        labelintersectionaz_diheading.pack()
        labelintersectionaz_di1.pack()
        labelintersectionaz_diimage1.pack()
        labelintersectionaz_di2.pack()
        labelintersectionaz_diimage2.pack()
        frame1.place(x=0,y=0,relheight=1.0,relwidth=1.0)
        f_intersectionaz_di.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        canvasintersectionaz_di.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        manageEnterWindow  
        
        
    elif item_iid == "I010":
        # ===========================Intersectiondi_di User  Guide==================================================
        frame1 = Frame(mainframe2)
        f_intersectiondi_di = Frame(frame1, bg='blue')
        canvasintersectiondi_di = Canvas(f_intersectiondi_di, borderwidth=0)
        frameintersectiondi_di00 = Frame(canvasintersectiondi_di)
        
        def mousewheel(event):canvasintersectiondi_di.yview_scroll(int(-1*(event.delta/120)),"units")
        def bind_to_wheel(event):canvasintersectiondi_di.bind_all("<MouseWheel>",mousewheel)
        def unbind_to_wheel(event):canvasintersectiondi_di.unbind_all("<MouseWheel>")
        f_intersectiondi_di.bind('<Enter>',bind_to_wheel)    
        f_intersectiondi_di.bind('<Leave>',unbind_to_wheel)  
        
        
        vsb = Scrollbar(f_intersectiondi_di, orient="vertical", command=canvasintersectiondi_di.yview)
        vsb.pack(side="right", fill="y")
        hsb = Scrollbar(f_intersectiondi_di, orient="horizontal", command=canvasintersectiondi_di.xview)
        hsb.pack(side="bottom", fill="x")
        #canvasintersectiondi_di.pack(side="left", fill="both", expand=True)
        canvasintersectiondi_di.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        
        frameintersectiondi_di00.bind("<Configure>", lambda event, canvas=canvasintersectiondi_di: onFrameConfigure(canvasintersectiondi_di))
        canvasintersectiondi_di.create_window((1,1), window=frameintersectiondi_di00, anchor="nw")
        canvasintersectiondi_di.configure(yscrollcommand=vsb.set)
        canvasintersectiondi_di.configure(xscrollcommand=hsb.set)
        
        
        imageintersectiondi_di1 = PhotoImage(file='Intersectiondi_di user guide_001.png')
        imageintersectiondi_di2 = PhotoImage(file='Intersectiondi_di user guide_002.png')
        imageintersectiondi_di3 = PhotoImage(file='Intersectiondi_di user guide_003.png')
        
        labelintersectiondi_di1 = Label(frameintersectiondi_di00, text='''
        This sub-module can solve the distance-distance intersection problem of position fixing, given the 
        coordinates of two control stations and the azimuth of one intersecting line and the distance of 
        the other.
 
        The program accepts direct input of all data necessary for the computation.
        
        The labels of the control stations are assumed “Control_A” and “Control_B” while the unknown station 
        label is assumed “Point_O”
        
        The data required from the user includes the following:
            
           1. Control_A: The coordinate value of the first control point.
           
           2. Control_B: The coordinate value of the second control point.
           
           3. Distance AO: The distance of the line from the first control point to the unknown point
              
           4. Distance BO: The distance of the line from the second control point to the unknown point
           
        **Note that the user is required to provide the clockwise order of the triangle formed by the controls 
          and the unknown point.

        While filling the data fields to be used by this sub-module, the user should note the following:
            
           1. All coordinate value fields are numeric, entered in easting and northing pairs, and separated by a
              comma, For example,  123456.789,987654.321.
              
           2. All distance fields are numeric. For example, input 123.456m as 123.456
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        # =============================================================================
        labelintersectiondi_diimage1 = Label(frameintersectiondi_di00, image=imageintersectiondi_di1)
        
        labelintersectiondi_di2 = Label(frameintersectiondi_di00, text='''
        Before computing the intersection, the user can check the box to 'Plot the Intersected Point'. Checking the 
        box provides a visual plot of the intersected point relative to the control stations.
        
        Finally, clicking the “COMPUTE” button does the magic.
        
        Upon computation, the results can be exported to different file formats as below:
            
           1. Shapefile (.shp): This export functionality can be done from the “File” menu” » “Export output” » “Export
              to shp”. It exports the plot of the intersected point relative to the control stations to a shapefile 
              which can be opened by any Geographic Information System (GIS) based software example ArcGIS and QGIS.
              
           2. Drawing Exchange Format (.dxf): This export functionality can be done from the “File” menu » “Export output” »
              “Export to dxf”. It exports the plot of the intersected point relative to the control stations to a drawing 
              exchange file format (.dxf) which can be opened by any Computer Aided Design (CAD) based software example 
              AutoCAD and IntelliCAD.
              
           3. Text file (.txt): This export functionality can be done from the “File” menu »  “Export output” » ”Export to
              txt”. It exports the generated coordinate of the unknown point and the additional information in the report 
              to text file format (.txt) which can be opened by any text-editing software example Notepad.
        
           4. Raster Image File Formats: This export functionality is only available if the “Plot the Intersected Point” 
              checkbox was checked before the reduction, that is, it can only be accessed from the “Plot the Intersected 
              Point” window. This export can be done from the “Plot the Intersected Point” toolbar » “Save the figure ( )”.
              Different image file formats are achievable example png, jpeg, pdf, tif, svg, pgf and so on.
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        labelintersectiondi_diimage2 = Label(frameintersectiondi_di00, image=imageintersectiondi_di2)
        labelintersectiondi_diheading = Label(frameintersectiondi_di00, image=imageintersectiondi_di3)
        
        labelintersectiondi_diheading.pack()
        labelintersectiondi_di1.pack()
        labelintersectiondi_diimage1.pack()
        labelintersectiondi_di2.pack()
        labelintersectiondi_diimage2.pack()
        frame1.place(x=0,y=0,relheight=1.0,relwidth=1.0)
        f_intersectiondi_di.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        canvasintersectiondi_di.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        manageEnterWindow  
        
        
    elif item_iid == "I011":
        # ===========================Intersectionaa_aa User  Guide==================================================
        frame1 = Frame(mainframe2)
        f_intersectionaa_aa = Frame(frame1, bg='blue')
        canvasintersectionaa_aa = Canvas(f_intersectionaa_aa, borderwidth=0)
        frameintersectionaa_aa00 = Frame(canvasintersectionaa_aa)
        
        def mousewheel(event):canvasintersectionaa_aa.yview_scroll(int(-1*(event.delta/120)),"units")
        def bind_to_wheel(event):canvasintersectionaa_aa.bind_all("<MouseWheel>",mousewheel)
        def unbind_to_wheel(event):canvasintersectionaa_aa.unbind_all("<MouseWheel>")
        f_intersectionaa_aa.bind('<Enter>',bind_to_wheel)    
        f_intersectionaa_aa.bind('<Leave>',unbind_to_wheel)  
        
        
        vsb = Scrollbar(f_intersectionaa_aa, orient="vertical", command=canvasintersectionaa_aa.yview)
        vsb.pack(side="right", fill="y")
        hsb = Scrollbar(f_intersectionaa_aa, orient="horizontal", command=canvasintersectionaa_aa.xview)
        hsb.pack(side="bottom", fill="x")
        #canvasintersectionaa_aa.pack(side="left", fill="both", expand=True)
        canvasintersectionaa_aa.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        
        frameintersectionaa_aa00.bind("<Configure>", lambda event, canvas=canvasintersectionaa_aa: onFrameConfigure(canvasintersectionaa_aa))
        canvasintersectionaa_aa.create_window((1,1), window=frameintersectionaa_aa00, anchor="nw")
        canvasintersectionaa_aa.configure(yscrollcommand=vsb.set)
        canvasintersectionaa_aa.configure(xscrollcommand=hsb.set)
        
        
        imageintersectionaa_aa1 = PhotoImage(file='Intersectionaa_aa user guide_001.png')
        imageintersectionaa_aa2 = PhotoImage(file='Intersectionaa_aa user guide_002.png')
        imageintersectionaa_aa3 = PhotoImage(file='Intersectionaa_aa user guide_003.png')
        
        labelintersectionaa_aa1 = Label(frameintersectionaa_aa00, text='''
        This sub-module can solve the distance-distance intersection problem of position fixing, given the 
        coordinates of two control stations and the interior angles observed at the control stations.
 
        The program accepts direct input of all data necessary for the computation.
        
        The labels of the control stations are assumed “Control_A” and “Control_B” while the unknown station 
        label is assumed “Point_O”
        
        The data required from the user includes the following:
            
           1. Control_A: The coordinate value of the first control point.
           
           2. Control_B: The coordinate value of the second control point.
           
           3. Angle BÂO: The interior angle subtended at the first control by the unknown point and the 
              second control.
              
           4. Angle AḂO: The interior angle subtended at the second control by the unknown point and the 
              first control.
           
        **Note that the user is required to provide the clockwise order of the triangle formed by the controls 
          and the unknown point.

        While filling the data fields to be used by this sub-module, the user should note the following:
            
           1. All coordinate value fields are numeric, entered in easting and northing pairs, and separated by a
              comma, For example,  123456.789,987654.321.
              
           2. All angular observation fields are numeric and in the form “dd mm ss”, with the degree, minutes and
              seconds values separated by space. For example, input 342°09'58" as 342 09 58
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        # =============================================================================
        labelintersectionaa_aaimage1 = Label(frameintersectionaa_aa00, image=imageintersectionaa_aa1)
        
        labelintersectionaa_aa2 = Label(frameintersectionaa_aa00, text='''
        Before computing the intersection, the user can check the box to 'Plot the Intersected Point'. Checking the 
        box provides a visual plot of the intersected point relative to the control stations.
        
        Finally, clicking the “COMPUTE” button does the magic.
        
        Upon computation, the results can be exported to different file formats as below:
            
           1. Shapefile (.shp): This export functionality can be done from the “File” menu” » “Export output” » “Export
              to shp”. It exports the plot of the intersected point relative to the control stations to a shapefile 
              which can be opened by any Geographic Information System (GIS) based software example ArcGIS and QGIS.
              
           2. Drawing Exchange Format (.dxf): This export functionality can be done from the “File” menu » “Export output” »
              “Export to dxf”. It exports the plot of the intersected point relative to the control stations to a drawing 
              exchange file format (.dxf) which can be opened by any Computer Aided Design (CAD) based software example 
              AutoCAD and IntelliCAD.
              
           3. Text file (.txt): This export functionality can be done from the “File” menu »  “Export output” » ”Export to
              txt”. It exports the generated coordinate of the unknown point and the additional information in the report 
              to text file format (.txt) which can be opened by any text-editing software example Notepad.
        
           4. Raster Image File Formats: This export functionality is only available if the “Plot the Intersected Point” 
              checkbox was checked before the reduction, that is, it can only be accessed from the “Plot the Intersected 
              Point” window. This export can be done from the “Plot the Intersected Point” toolbar » “Save the figure ( )”.
              Different image file formats are achievable example png, jpeg, pdf, tif, svg, pgf and so on.
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        labelintersectionaa_aaimage2 = Label(frameintersectionaa_aa00, image=imageintersectionaa_aa2)
        labelintersectionaa_aaheading = Label(frameintersectionaa_aa00, image=imageintersectionaa_aa3)
        
        labelintersectionaa_aaheading.pack()
        labelintersectionaa_aa1.pack()
        labelintersectionaa_aaimage1.pack()
        labelintersectionaa_aa2.pack()
        labelintersectionaa_aaimage2.pack()
        frame1.place(x=0,y=0,relheight=1.0,relwidth=1.0)
        f_intersectionaa_aa.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        canvasintersectionaa_aa.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        manageEnterWindow  


    elif item_iid == "I012":
        # ===========================Intersectionaa_sd User  Guide==================================================
        frame1 = Frame(mainframe2)
        f_intersectionaa_sd = Frame(frame1, bg='blue')
        canvasintersectionaa_sd = Canvas(f_intersectionaa_sd, borderwidth=0)
        frameintersectionaa_sd00 = Frame(canvasintersectionaa_sd)
        
        def mousewheel(event):canvasintersectionaa_sd.yview_scroll(int(-1*(event.delta/120)),"units")
        def bind_to_wheel(event):canvasintersectionaa_sd.bind_all("<MouseWheel>",mousewheel)
        def unbind_to_wheel(event):canvasintersectionaa_sd.unbind_all("<MouseWheel>")
        f_intersectionaa_sd.bind('<Enter>',bind_to_wheel)    
        f_intersectionaa_sd.bind('<Leave>',unbind_to_wheel)  
        
        
        vsb = Scrollbar(f_intersectionaa_sd, orient="vertical", command=canvasintersectionaa_sd.yview)
        vsb.pack(side="right", fill="y")
        hsb = Scrollbar(f_intersectionaa_sd, orient="horizontal", command=canvasintersectionaa_sd.xview)
        hsb.pack(side="bottom", fill="x")
        #canvasintersectionaa_sd.pack(side="left", fill="both", expand=True)
        canvasintersectionaa_sd.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        
        frameintersectionaa_sd00.bind("<Configure>", lambda event, canvas=canvasintersectionaa_sd: onFrameConfigure(canvasintersectionaa_sd))
        canvasintersectionaa_sd.create_window((1,1), window=frameintersectionaa_sd00, anchor="nw")
        canvasintersectionaa_sd.configure(yscrollcommand=vsb.set)
        canvasintersectionaa_sd.configure(xscrollcommand=hsb.set)
        
        
        imageintersectionaa_sd1 = PhotoImage(file='Intersectionaa_sd user guide_001.png')
        imageintersectionaa_sd2 = PhotoImage(file='Intersectionaa_sd user guide_002.png')
        imageintersectionaa_sd3 = PhotoImage(file='Intersectionaa_sd user guide_003.png')
        
        labelintersectionaa_sd1 = Label(frameintersectionaa_sd00, text='''
        This sub-module can solve a special intersection problem of position fixing, where the second control is 
        inaccessible (Side section). In the method of side section using inner angles, the angle subtended at one
        of the controls is provided and due to the inaccessibility of the second control, the angle subtended at 
        the unknown point by the controls is observed and used as compensation.  Also given are the coordinates of
        two control stations. 
 
        The program accepts direct input of all data necessary for the computation.
        
        The labels of the control stations are assumed “Control_A” and “Control_B” while the unknown station 
        label is assumed “Point_O”
        
        **Note that the control station at which the interior angle is observed is considered Control_A 
          (first control)
          
        **Note that the control station at which the interior angle is not observed due to inaccessibility is 
          considered Control_B (second control)

        The data required from the user includes the following:
            
           1. Control_A: The coordinate value of the first control point.
           
           2. Control_B: The coordinate value of the second control point.
           
           3. Angle BÂO: The interior angle subtended at the first control by the unknown point and the 
              second control.
              
           4. Angle AÔB: The interior angle subtended at the unknown point by the two controls 
           
        **Note that the user is required to provide the clockwise order of the triangle formed by the controls 
          and the unknown point.

        While filling the data fields to be used by this sub-module, the user should note the following:
            
           1. All coordinate value fields are numeric, entered in easting and northing pairs, and separated by a
              comma, For example,  123456.789,987654.321.
              
           2. All angular observation fields are numeric and in the form “dd mm ss”, with the degree, minutes and
              seconds values separated by space. For example, input 342°09'58" as 342 09 58
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        # =============================================================================
        labelintersectionaa_sdimage1 = Label(frameintersectionaa_sd00, image=imageintersectionaa_sd1)
        
        labelintersectionaa_sd2 = Label(frameintersectionaa_sd00, text='''
        Before computing the intersection, the user can check the box to 'Plot the Intersected Point'. Checking the 
        box provides a visual plot of the intersected point relative to the control stations.
        
        Finally, clicking the “COMPUTE” button does the magic.
        
        Upon computation, the results can be exported to different file formats as below:
            
           1. Shapefile (.shp): This export functionality can be done from the “File” menu” » “Export output” » “Export
              to shp”. It exports the plot of the intersected point relative to the control stations to a shapefile 
              which can be opened by any Geographic Information System (GIS) based software example ArcGIS and QGIS.
              
           2. Drawing Exchange Format (.dxf): This export functionality can be done from the “File” menu » “Export output” »
              “Export to dxf”. It exports the plot of the intersected point relative to the control stations to a drawing 
              exchange file format (.dxf) which can be opened by any Computer Aided Design (CAD) based software example 
              AutoCAD and IntelliCAD.
              
           3. Text file (.txt): This export functionality can be done from the “File” menu »  “Export output” » ”Export to
              txt”. It exports the generated coordinate of the unknown point and the additional information in the report 
              to text file format (.txt) which can be opened by any text-editing software example Notepad.
        
           4. Raster Image File Formats: This export functionality is only available if the “Plot the Intersected Point” 
              checkbox was checked before the reduction, that is, it can only be accessed from the “Plot the Intersected 
              Point” window. This export can be done from the “Plot the Intersected Point” toolbar » “Save the figure ( )”.
              Different image file formats are achievable example png, jpeg, pdf, tif, svg, pgf and so on.
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        labelintersectionaa_sdimage2 = Label(frameintersectionaa_sd00, image=imageintersectionaa_sd2)
        labelintersectionaa_sdheading = Label(frameintersectionaa_sd00, image=imageintersectionaa_sd3)
        
        labelintersectionaa_sdheading.pack()
        labelintersectionaa_sd1.pack()
        labelintersectionaa_sdimage1.pack()
        labelintersectionaa_sd2.pack()
        labelintersectionaa_sdimage2.pack()
        frame1.place(x=0,y=0,relheight=1.0,relwidth=1.0)
        f_intersectionaa_sd.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        canvasintersectionaa_sd.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        manageEnterWindow 
        
        
    elif item_iid == "I013":
        # ===========================Intersectiondi_sd User  Guide==================================================
        frame1 = Frame(mainframe2)
        f_intersectiondi_sd = Frame(frame1, bg='blue')
        canvasintersectiondi_sd = Canvas(f_intersectiondi_sd, borderwidth=0)
        frameintersectiondi_sd00 = Frame(canvasintersectiondi_sd)
        
        def mousewheel(event):canvasintersectiondi_sd.yview_scroll(int(-1*(event.delta/120)),"units")
        def bind_to_wheel(event):canvasintersectiondi_sd.bind_all("<MouseWheel>",mousewheel)
        def unbind_to_wheel(event):canvasintersectiondi_sd.unbind_all("<MouseWheel>")
        f_intersectiondi_sd.bind('<Enter>',bind_to_wheel)    
        f_intersectiondi_sd.bind('<Leave>',unbind_to_wheel)  
        
        
        vsb = Scrollbar(f_intersectiondi_sd, orient="vertical", command=canvasintersectiondi_sd.yview)
        vsb.pack(side="right", fill="y")
        hsb = Scrollbar(f_intersectiondi_sd, orient="horizontal", command=canvasintersectiondi_sd.xview)
        hsb.pack(side="bottom", fill="x")
        #canvasintersectiondi_sd.pack(side="left", fill="both", expand=True)
        canvasintersectiondi_sd.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        
        frameintersectiondi_sd00.bind("<Configure>", lambda event, canvas=canvasintersectiondi_sd: onFrameConfigure(canvasintersectiondi_sd))
        canvasintersectiondi_sd.create_window((1,1), window=frameintersectiondi_sd00, anchor="nw")
        canvasintersectiondi_sd.configure(yscrollcommand=vsb.set)
        canvasintersectiondi_sd.configure(xscrollcommand=hsb.set)
        
        
        imageintersectiondi_sd1 = PhotoImage(file='Intersectiondi_sd user guide_001.png')
        imageintersectiondi_sd2 = PhotoImage(file='Intersectiondi_sd user guide_002.png')
        imageintersectiondi_sd3 = PhotoImage(file='Intersectiondi_sd user guide_003.png')
        
        labelintersectiondi_sd1 = Label(frameintersectiondi_sd00, text='''
        This sub-module can solve a special intersection problem of position fixing, where the second control 
        is inaccessible (Side section). In the method of side section using distance, the angle subtended at 
        one of the controls is provided and due to the inaccessibility of the second control, the distance from
        one of the controls to the unknown point is observed and used as compensation.  Also given are the 
        coordinates of two control stations.  
 
        The program accepts direct input of all data necessary for the computation.
        
        The labels of the control stations are assumed “Control_A” and “Control_B” while the unknown station 
        label is assumed “Point_O”
        
        **Note that the control station for which the distance to the unknown point is observed is considered
          Control_A (first control)
          
        **Note that the control station for which the distance to the unknown point is not observed due to 
          inaccessibility is  considered Control_B (second control)

        The data required from the user includes the following:
            
           1. Control_A: The coordinate value of the first control point.
           
           2. Control_B: The coordinate value of the second control point.
           
           3. Angle AÔB: The interior angle subtended at the unknown point by the two controls 
              
           4. Distance AO: The distance between the first control and the unknown point.
           
        **Note that the user is required to provide the clockwise order of the triangle formed by the controls 
          and the unknown point.

        While filling the data fields to be used by this sub-module, the user should note the following:
            
           1. All coordinate value fields are numeric, entered in easting and northing pairs, and separated by a
              comma, For example,  123456.789,987654.321.
              
           2. All distance fields are numeric. For example, input 123.456m as 123.456
           
           3. All angular observation fields are numeric and in the form “dd mm ss”, with the degree, minutes and
              seconds values separated by space. For example, input 342°09'58" as 342 09 58
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        # =============================================================================
        labelintersectiondi_sdimage1 = Label(frameintersectiondi_sd00, image=imageintersectiondi_sd1)
        
        labelintersectiondi_sd2 = Label(frameintersectiondi_sd00, text='''
        Before computing the intersection, the user can check the box to 'Plot the Intersected Point'. Checking the 
        box provides a visual plot of the intersected point relative to the control stations.
        
        Finally, clicking the “COMPUTE” button does the magic.
        
        Upon computation, the results can be exported to different file formats as below:
            
           1. Shapefile (.shp): This export functionality can be done from the “File” menu” » “Export output” » “Export
              to shp”. It exports the plot of the intersected point relative to the control stations to a shapefile 
              which can be opened by any Geographic Information System (GIS) based software example ArcGIS and QGIS.
              
           2. Drawing Exchange Format (.dxf): This export functionality can be done from the “File” menu » “Export output” »
              “Export to dxf”. It exports the plot of the intersected point relative to the control stations to a drawing 
              exchange file format (.dxf) which can be opened by any Computer Aided Design (CAD) based software example 
              AutoCAD and IntelliCAD.
              
           3. Text file (.txt): This export functionality can be done from the “File” menu »  “Export output” » ”Export to
              txt”. It exports the generated coordinate of the unknown point and the additional information in the report 
              to text file format (.txt) which can be opened by any text-editing software example Notepad.
        
           4. Raster Image File Formats: This export functionality is only available if the “Plot the Intersected Point” 
              checkbox was checked before the reduction, that is, it can only be accessed from the “Plot the Intersected 
              Point” window. This export can be done from the “Plot the Intersected Point” toolbar » “Save the figure ( )”.
              Different image file formats are achievable example png, jpeg, pdf, tif, svg, pgf and so on.
        ''',justify='left',font=('helvetica', 10,'italic'))
        
        labelintersectiondi_sdimage2 = Label(frameintersectiondi_sd00, image=imageintersectiondi_sd2)
        labelintersectiondi_sdheading = Label(frameintersectiondi_sd00, image=imageintersectiondi_sd3)
        
        labelintersectiondi_sdheading.pack()
        labelintersectiondi_sd1.pack()
        labelintersectiondi_sdimage1.pack()
        labelintersectiondi_sd2.pack()
        labelintersectiondi_sdimage2.pack()
        frame1.place(x=0,y=0,relheight=1.0,relwidth=1.0)
        f_intersectiondi_sd.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        canvasintersectiondi_sd.place(x=1,y=1,relheight=1.0,relwidth=1.0)
        manageEnterWindow 

style = ttk.Style()
style.configure("mystyle.Treeview", bd=10, font=(None,11),rowheight=30)
style.configure("mystyle.Treeview.Heading",font=('Calibri',13,'bold'))
style.layout("mystyle.Treeview",[('mystyle.Treeview.treearea',{'sticky': 'nswe'})])

mainframe1 = tk.Frame(panedwindow)#)
coordinates=ttk.Treeview(mainframe1, style="mystyle.Treeview", height='30')
ybar = Scrollbar(mainframe1,orient=VERTICAL, command=coordinates.yview)
coordinates.configure(yscroll=ybar.set)
#ybar.place(relx=1.0,rely=0,relheight=1.0,anchor=NE)
ybar.pack(side="right", fill="y")


xbar = Scrollbar(mainframe1,orient=HORIZONTAL, command=coordinates.xview)
coordinates.configure(xscroll=xbar.set)
#xbar.place(relx=0,rely=1.0,relwidth=1.0,anchor=SW)
xbar.pack(side="bottom", fill="x")


coordinates.pack(side=LEFT, fill=BOTH, expand=True)
coordinates.column('#0',width=180,minwidth=260,stretch=YES)
coordinates.heading('#0',text = "User Guide",anchor=W)                  
coordinates.bind("<FocusIn>",userguides)
coordinates.bind("<ButtonRelease-1>",userguides)
                    
panedwindow.add(mainframe1)
panedwindow.add(mainframe2)

aboutfolder = coordinates.insert('', 1,text='About')
menufolder = coordinates.insert('', 2,text='Menus')
programfolder = coordinates.insert('', 3,text='Programs')


filemenu = coordinates.insert(menufolder, 4,text='File Menu')
traversemenu = coordinates.insert(menufolder, 5,text='Traverse Menu')

traversefolder = coordinates.insert(programfolder, "end",text='Traverse Computation')
levelfolder = coordinates.insert(programfolder, "end",text='Level Computation')
hydrofolder = coordinates.insert(programfolder, "end",text='Sounding Reduction')

areafolder = coordinates.insert(programfolder, "end",text='Area Computation')
coordinates.insert(areafolder, "end",text='Reading Coordinates from File')
coordinates.insert(areafolder, "end",text='Inputting coordinates directly')

resectionfolder = coordinates.insert(programfolder, "end",text='Three Point Resection')
intersectionfolder = coordinates.insert(programfolder, "end",text='Intersection')
coordinates.insert(intersectionfolder, "end",text='Azimuth-Azimuth Intersection')
coordinates.insert(intersectionfolder, "end",text='Azimuth_Distance Intersection')
coordinates.insert(intersectionfolder, "end",text='Distance_Distance Intersection')
coordinates.insert(intersectionfolder, "end",text='Angle_Angle Intersection')
coordinates.insert(intersectionfolder, "end",text='Interior Angle Side section')
coordinates.insert(intersectionfolder, "end",text='Distance Side section')


levelmenu = coordinates.insert(menufolder,"end",text='Level Menu')
bathymetrymenu = coordinates.insert(menufolder,"end",text='Bathymetry Menu')
areamenu = coordinates.insert(menufolder,"end",text='Area Menu')
resectionmenu = coordinates.insert(menufolder,"end",text='Resection Menu')
intersectionmenu = coordinates.insert(menufolder,"end",text='Intersection Menu')
helpmenu = coordinates.insert(menufolder, "end",text='Help Menu')


#root.mainloop()       