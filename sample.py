import tkinter as tk
from tkinter import *
from tkinter import filedialog, StringVar, Frame, OptionMenu, Label, Canvas
import pandas as pd

import_file_path = None
root = tk.Tk()
root.title("Pricing Generator")

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
    global xls
    global import_file_path

    import_file_path = filedialog.askopenfilename()
    df = pd.read_excel(import_file_path)
    xls = pd.ExcelFile(import_file_path)
    choices = xls.sheet_names
    print(choices)
    popupMenu = OptionMenu(mainframe, tkvar, *choices)
    Label(mainframe, text="Choose a Product").grid(row = 2, column = 1)
    popupMenu.grid(row = 3, column =1)

browseButton_Excel = Button(mainframe, text='Import Excel File', command=getExcel, fg="white")
browseButton_Excel.grid(row = 1, column = 1)

# on change dropdown value
def change_dropdown(*args):
    global prod_list
    rows_to_skip = list(range(1,51))
    print( tkvar.get() )
    sheet = tkvar.get()
    Label(mainframe, text=tkvar.get()).grid(row=3, column=2)
    prod_list = pd.read_excel(import_file_path, sheet_name=sheet, usecols = "A:K")
    row_index = prod_list.iloc[:, 1].str.match('UNIT').index
    print (row_index)
    print (prod_list)

# link function to change dropdown
tkvar.trace('w', change_dropdown)

def getList(dict):
    return dict.keys()

root.mainloop()