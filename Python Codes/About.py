# -*- coding: utf-8 -*-
"""
Created on Sun Dec 20 01:52:53 2020

@author: IFEANYI
"""

from tkinter import *
from tkinter import ttk

rootabout = Toplevel()
rootabout.resizable(width=False, height=False)
rootabout.title('Survey Companion: About') 

icon = PhotoImage(file = 'S_Companion Icon.png')
rootabout.iconphoto(False,icon)

frameabouthome = Frame(rootabout)
imageabouthome1 = PhotoImage(file='Abouthome user guide_001.png')

labelabouthome1 = Label(frameabouthome, text='''
Designed and developed by UGWUANYI, IFEANYI ALEXANDER

    Tel.: +234 814 138 3318
    Email: ugwuanyialexifyco@gmail.com
                 ifeanyi.ugwuanyi.197580@unn.edu.ng

''',justify='left',font=('helvetica', 12,'italic','bold'))

labelabouthomeimage1 = Label(frameabouthome, image=imageabouthome1)
labelabouthomeimage1.pack()
labelabouthome1.pack()
frameabouthome.pack()



