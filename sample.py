import tkinter as tk
import pandas as pd
from tkinter import *
from tkinter import filedialog, StringVar, Frame, OptionMenu, Label, Entry
from tkinter import ttk


class ProposalApp(Frame):

    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):

        import_file_path = None
        self.title("BoM Generator")
        global item_row_num

        # Add a grid
        mainframe = Frame(self)
        mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
        mainframe.columnconfigure(0, weight=1)
        mainframe.rowconfigure(0, weight=1)
        mainframe.pack(pady=100, padx=100)
        # Create a Tkinter variable
        tkvar = StringVar(self)
        tree = ttk.Treeview(mainframe)
        browseButton_Excel = ttk.Button(mainframe, text='Import Price Sheet', command=self.getExcel)
        browseButton_Excel.grid(row=1, column=0)
        # link function to change dropdown
        tkvar.trace('w', self.change_dropdown)


    def getExcel(self):
        global df
        global xls
        global import_file_path

        import_file_path = filedialog.askopenfilename()
        df = pd.read_excel(import_file_path)
        xls = pd.ExcelFile(import_file_path)
        choices = xls.sheet_names
        print('Sheet Names:', choices)
        popupMenu = OptionMenu(self.mainframe, self.tkvar, *choices)
        Label(self.mainframe, text="Choose a Product").grid(row = 2, column = 0)
        popupMenu.grid(row = 3, column =0)

    # on change dropdown value
    def change_dropdown(self, *args):
        global prod_list
        global tree
        print( 'Selected Sheet Name:', self.tkvar.get() )
        sheet = self.tkvar.get()
        self.tkvar.set(self.tkvar.get())
        Label(self.mainframe, text=self.tkvar.get()).grid(row=3, column=0)
        prod_list = pd.read_excel(import_file_path, sheet_name=sheet, usecols = "A:K")
        print ('Original Columns:', prod_list.columns)
        prod_list.fillna(0, inplace=True)
        prod_list.rename(columns = {'Unnamed: 0':'Unit', 'Unnamed: 1':'SKU', 'Unnamed: 2':'Description',
                                'Unnamed: 3':'Price', 'Unnamed: 4':'oneyr', 'Unnamed: 5':'twoyr',
                                'Unnamed: 6':'threeyr', 'Unnamed: 7':'fouryr', 'Unnamed: 8':'fiveyr',
                                'Unnamed: 9':'Comments', 'Unnamed: 10':'Cat'}, inplace=True)
        print ('New Column Names:', prod_list.columns)
        start_index = self.getIndexes(prod_list, 'UNIT')
        idx, val = start_index[0]
        print ('Starting Index:', idx)
        prod_list.drop(prod_list.index[:idx], inplace=True)
        self.getUnits(prod_list, idx)
        tree["columns"] = ("SKU", "Description", "Price", "1 Year", "2 Year", "3 Year", "4 Year", "5 Year","Comments", "Category")
        tree.column("SKU", width=400, minwidth=100)
        tree.column("Description", width=80, minwidth=50, stretch=tk.NO)
        tree.column("Price", width=20, minwidth=10)
        tree.column("1 Year", width=20, minwidth=10)
        tree.column("2 Year", width=20, minwidth=10)
        tree.column("3 Year", width=20, minwidth=10)
        tree.column("4 Year", width=20, minwidth=10)
        tree.column("5 Year", width=20, minwidth=10)
        tree.column("Comments", width=100, minwidth=50)
        tree.column("Category", width=10, minwidth=5)
        tree.heading("#0",text="",anchor=tk.W)
        tree.heading("SKU", text="SKU",anchor=tk.W)
        tree.heading("Description", text="Description",anchor=tk.W)
        tree.heading("Price", text="Price", anchor=tk.W)
        tree.heading("1 Year", text="1 Year", anchor=tk.W)
        tree.heading("2 Year", text="2 Year", anchor=tk.W)
        tree.heading("3 Year", text="3 Year", anchor=tk.W)
        tree.heading("4 Year", text="4 Year", anchor=tk.W)
        tree.heading("5 Year", text="5 Year", anchor=tk.W)
        tree.heading("Comments", text="Comments", anchor=tk.W)
        tree.heading("Category", text="Category", anchor=tk.W)
        tree.grid(row=4, column=0, sticky=tk.S + tk.W + tk.E + tk.N)

    def getIndexes(dfObj, value):
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
        curItems = tree.selection()
        for i in curItems:
            entry = tree.item(i)['values']
            print(entry)
            global item_row_num
            item_row_num +=1
            item_row = tk.Label(self.outframe, text=entry)
            item_row.grid(column=0, row=item_row_num)
            item_row.pack

    def getUnits(self, data, index):
        print (len(data))
        index = int(index)
        unit_elements = {}
        unit = ""
        folder = 'folder'
        header_row = False
        counter = 0
        global item_row_num
        item_row_num = 0
        for row in data.itertuples():
            # access data using column names
            if (getattr(row, 'Unit') == 'UNIT') and (getattr(row, 'SKU') == 'SKU'):
                header_row = True
                continue
            else:
                if header_row:
                    header_row = False
                    unit = getattr(row, 'Unit')
                    unit_elements.setdefault(unit, {})
                    parent_id=counter
                    tree.insert(parent='', index='end', iid=parent_id, text=unit, values=("", "", "", "", "", "", "", "", ""))
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
                counter +=1
                tree.insert(parent=parent_id, index='end', iid=counter, values=(sku,descr,price,oneyr,twoyr,
                                    threeyr,fouryr,fiveyr,comments,cat,))
                unit_elements[unit].update({'sku': sku, 'descr': descr, 'price': price, 'oneyr': oneyr, 'twoyr': twoyr,
                                    'threeyr': threeyr, 'fouryr': fouryr, 'fiveyr': fiveyr, 'comments': comments,
                                    'cat': cat})
            counter +=1
        tree.pack(side=tk.TOP, fill=tk.X)
        tree.bind('<Double-1>', lambda e: self.select())
        addButton = ttk.Button(self.mainframe, text='Add', command=self.select())
        addButton.grid(row=5, column=0)
        outframe = Frame(self)
        outframe.grid(column=0, row=6, sticky=(N, W, E, S))
        outframe.columnconfigure(0, weight=1)
        outframe.rowconfigure(0, weight=1)
        outframe.pack(pady=100, padx=300)

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
    app = ProposalApp
    root.mainloop()

if __name__ == '__main__':
    main()
