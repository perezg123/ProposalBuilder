import tkinter as tk
from tkinter import *
from tkinter import filedialog, StringVar, Frame, OptionMenu, Label, Canvas
import pandas as pd

import_file_path = None
root = tk.Tk()
root.title("Pricing Generator")

#canvas1.create_window(150, 150, window=browseButton_Excel)
#canvas1 = tk.Canvas(root, width=300, height=300)
#canvas1.pack()
# Add a grid
mainframe = Frame(root)
mainframe.grid(column=0,row=0, sticky=(N,W,E,S) )
mainframe.columnconfigure(0, weight = 1)
mainframe.rowconfigure(0, weight = 1)
mainframe.pack(pady = 100, padx = 100)
# Create a Tkinter variable
tkvar = StringVar(root)

def getExcel():
    global df
    global import_file_path

    import_file_path = filedialog.askopenfilename()
    df = pd.read_excel(import_file_path)
    xls = pd.ExcelFile(import_file_path)
    print(xls.sheet_names)
    choices = xls.sheet_names
    print(choices)
    popupMenu = OptionMenu(mainframe, tkvar, *choices)
    Label(mainframe, text="Choose a Product").grid(row = 2, column = 1)
    popupMenu.grid(row = 3, column =1)


browseButton_Excel = tk.Button(mainframe, text='Import Excel File', command=getExcel, fg="black")
browseButton_Excel.grid(row = 1, column = 1)

# on change dropdown value
def change_dropdown(*args):
    print( tkvar.get() )
    # link function to change dropdown
    tkvar.trace('w', change_dropdown)

def getList(dict):
    return dict.keys()

#import_file_path = filedialog.askopenfilename()

root.mainloop()