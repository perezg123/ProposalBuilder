import tkinter as tk
import pandas as pd
from tkinter import *
from tkinter import filedialog, StringVar, Frame, OptionMenu, Label, Entry, Tk, ttk, Button, messagebox


class ProposalApp:

    global currSheet
    global df
    global xls
    global import_file_path
    global prod_list
    global prodTree
    global item_row_num

    def __init__(self, master):

        self.master = master
        import_file_path = None
        master.title("Materials List Creator")
        # Add a main grid with button, dropdown and tree
        self.mainframe = Frame(master)
        self.mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
        self.mainframe.columnconfigure(0, weight=1)
        self.mainframe.rowconfigure(0, weight=1)
        self.mainframe.pack(pady=100, padx=100)
        # Create a Tkinter variable
        self.currSheet = StringVar(master)
        self.prodTree = ttk.Treeview(self.mainframe)
        browseButton_Excel = ttk.Button(self.mainframe, text='Import Price Sheet', command=self.getExcel)
        browseButton_Excel.grid(row=1, column=0)
        self.currSheet.trace('w', self.change_dropdown)


    def getExcel(self):

        # Take an excel filepath, inspect the file, read the tabs and populate the "product" dropdown
        self.import_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        self.df = pd.read_excel(self.import_file_path)
        self.xls = pd.ExcelFile(self.import_file_path)
        choices = self.xls.sheet_names
        if ('Cover Sheet' or 'Index' or "General Info") not in choices:
            self.messagebox = messagebox
            self.title = "Input Error"
            self.message = "This does not appear to be a valid Price Sheet file."
            messagebox.showwarning(title=self.title, message=self.message)
            return
        print('Sheet Names:', choices)
        popupMenu = OptionMenu(self.mainframe, self.currSheet, *choices)
        Label(self.mainframe, text="Choose a Product").grid(row = 2, column = 0)
        popupMenu.config(width=30)
        popupMenu.grid(row = 2, column =1)

    def change_dropdown(self, *args):

        global item_row_num
        item_row_num = 0
        print( 'Selected Sheet Name:', self.currSheet.get() )
        sheet = self.currSheet.get()
        self.currSheet.set(self.currSheet.get())
        self.prod_list = pd.read_excel(self.import_file_path, sheet_name=sheet, usecols = "A:K")
        print ('Original Columns:', self.prod_list.columns)
        self.prod_list.fillna(0, inplace=True)
        self.prod_list.rename(columns = {'Unnamed: 0':'Unit', 'Unnamed: 1':'SKU', 'Unnamed: 2':'Description',
                                'Unnamed: 3':'Price', 'Unnamed: 4':'oneyr', 'Unnamed: 5':'twoyr',
                                'Unnamed: 6':'threeyr', 'Unnamed: 7':'fouryr', 'Unnamed: 8':'fiveyr',
                                'Unnamed: 9':'Comments', 'Unnamed: 10':'Cat'}, inplace=True)
        print ('New Column Names:', self.prod_list.columns)
        start_index = self.getIndexes(self.prod_list, 'UNIT')
        idx, val = start_index[0]
        print ('Starting Index:', idx)
        self.prod_list.drop(self.prod_list.index[:idx], inplace=True)
        self.getUnits(self.prod_list, idx)
        self.prodTree["columns"] = ("SKU", "Description", "Price", "1 Year", "2 Year", "3 Year", "4 Year",
                                    "5 Year","Comments", "Category")
        self.prodTree.column("SKU", width=400, minwidth=100)
        self.prodTree.column("Description", width=80, minwidth=50, stretch=tk.NO)
        self.prodTree.column("Price", width=20, minwidth=10)
        self.prodTree.column("1 Year", width=20, minwidth=10)
        self.prodTree.column("2 Year", width=20, minwidth=10)
        self.prodTree.column("3 Year", width=20, minwidth=10)
        self.prodTree.column("4 Year", width=20, minwidth=10)
        self.prodTree.column("5 Year", width=20, minwidth=10)
        self.prodTree.column("Comments", width=100, minwidth=50)
        self.prodTree.column("Category", width=10, minwidth=5)
        self.prodTree.heading("#0",text="",anchor=tk.W)
        self.prodTree.heading("SKU", text="SKU",anchor=tk.W)
        self.prodTree.heading("Description", text="Description",anchor=tk.W)
        self.prodTree.heading("Price", text="Price", anchor=tk.W)
        self.prodTree.heading("1 Year", text="1 Year", anchor=tk.W)
        self.prodTree.heading("2 Year", text="2 Year", anchor=tk.W)
        self.prodTree.heading("3 Year", text="3 Year", anchor=tk.W)
        self.prodTree.heading("4 Year", text="4 Year", anchor=tk.W)
        self.prodTree.heading("5 Year", text="5 Year", anchor=tk.W)
        self.prodTree.heading("Comments", text="Comments", anchor=tk.W)
        self.prodTree.heading("Category", text="Category", anchor=tk.W)
        self.prodTree.grid(row=4, column=0, sticky=tk.S + tk.W + tk.E + tk.N)

    def getIndexes(self, dfObj, value):
        ''' Get index positions of value in dataframe i.e. dfObj.'''

        listOfPos = list()
        # Get bool dataframe with True at positions where the given value exists
        result = dfObj.isin([value])
        # Get list of columns that contains the value
        seriesObj = result.any()
        columnNames = list(seriesObj[seriesObj == True].index)
        # Iterate over list of columns and fetch the rows indexes where value exists
        for col in columnNames:
            rows = list(result[col][result[col] == True].index)
            for row in rows:
                listOfPos.append((row, col))
        # Return a list of tuples indicating the positions of value in the dataframe
        return listOfPos

    def getList(dict):
        return dict.keys()

    def select(self):

        global item_row_num
        curItems = self.prodTree.selection()
        for i in curItems:
            entry = self.prodTree.item(i)['values']
            print(entry)
            item_row_num +=1
            self.item_row = tk.Label(self.outframe, text=entry)
            self.item_row.grid(column=0, row=item_row_num)
            self.item_row.pack

    def getUnits(self, data, index):

        print (len(data))
        self.index = int(index)
        self.unit_elements = {}
        self.unit = ""
        self.folder = 'folder'
        self.header_row = False
        self.counter = 0
        for row in data.itertuples():
            # access data using column names
            if (getattr(row, 'Unit') == 'UNIT') and (getattr(row, 'SKU') == 'SKU'):
                self.header_row = True
                continue
            else:
                if self.header_row:
                    self.header_row = False
                    self.unit = getattr(row, 'Unit')
                    self.unit_elements.setdefault(self.unit, {})
                    self.parent_id=self.counter
                    self.prodTree.insert(parent='', index='end', iid=self.parent_id, text=self.unit,
                                     values=("", "", "", "", "", "", "", "", ""))
                sku = getattr(row, 'SKU')
                descr = getattr(row, 'Description')
                float("{:.2f}".format(13.949999999999999))
                price = float("{:.2f}".format(getattr(row, 'Price')))
                oneyr = float("{:.2f}".format(getattr(row, 'oneyr')))
                twoyr = float("{:.2f}".format(getattr(row, 'twoyr')))
                threeyr = float("{:.2f}".format(getattr(row, 'threeyr')))
                fouryr = float("{:.2f}".format(getattr(row, 'fouryr')))
                fiveyr = float("{:.2f}".format(getattr(row, 'fiveyr')))
                comments = getattr(row, 'Comments')
                cat = getattr(row, 'Cat')
                self.counter +=1
                self.prodTree.insert(parent=self.parent_id, index='end', iid=self.counter,
                                 values=(sku,descr,price,oneyr,twoyr,threeyr,fouryr,fiveyr,comments,cat,))
                self.unit_elements[self.unit].update({'sku': sku, 'descr': descr, 'price': price, 'oneyr':
                                oneyr, 'twoyr': twoyr, 'threeyr': threeyr, 'fouryr': fouryr, 'fiveyr': fiveyr,
                                'comments': comments, 'cat': cat})
            self.counter +=1
        self.prodTree.pack(side=tk.TOP, fill=tk.X)
        self.prodTree.bind('<Double-1>', lambda e: self.select())
        self.outframe = Frame()
        self.outframe.grid(column=0, row=6, sticky=(N, W, E, S))
        self.outframe.columnconfigure(0, weight=1)
        self.outframe.rowconfigure(0, weight=1)
        self.outframe.pack(pady=100, padx=300)

    def convert_currency(val):
        """
        Convert the string number value to a float
         - Remove $
         - Remove commas
         - Convert to float type
        """
        new_val = val.replace(',','').replace('$', '')
        return float(new_val)

def main():

    root = tk.Tk()
    app = ProposalApp(root)
    root.mainloop()

if __name__ == '__main__':
    main()
