import logging
import os
import sys
from tkinter import *
from tkinter import filedialog, messagebox

import pandas as pd
from pandastable import Table, images
from pandastable.dialogs import addButton

from core.manager import PartsListsManager

VER = 'v0.1.2'


class App(Frame):
    """Podstawowe okno dla Frame"""

    def __init__(self, parent=None):
        self.parent = parent
        Frame.__init__(self)
        self.main = self.master
        self.main.geometry('1200x700')
        self.main.title('PartsListsManager ' + VER)

        self.toolbar = Frame(self.main)
        self.toolbar.pack(fill='both')

        f = Frame(self.main)
        f.pack(fill='both', expand=True)

        self.pt = MyTable(f, showtoolbar=False, showstatusbar=True)

        # - toolbar -
        self.create_toolbar()
        # - show -
        self.pt.show()

        return

    def create_toolbar(self):
        """Utworzenie toolbar
        """
        img = images.open_proj()
        addButton(self.toolbar,
                  'Open',
                  self.pt.load, img,
                  side='left',
                  tooltip='open PartsLists')

        img = images.save()
        addButton(self.toolbar,
                  'Load table',
                  self.pt.saveAs, img,
                  side='left',
                  tooltip='save as xlsx')

        img = images.cross()
        addButton(self.toolbar,
                  'Clear',
                  self.pt.clear_table, img,
                  side='left',
                  tooltip='clear table')

        img = images.aggregate()
        addButton(self.toolbar,
                  'Aggregate',
                  self.pt.aggregate, img,
                  side='left',
                  tooltip='aggregate')

        img = images.pivot()
        addButton(self.toolbar,
                  'Pivot',
                  self.pt.pivot, img,
                  side='left',
                  tooltip='pivot')


class MyTable(Table):
    """Modyfikacje klasy pandastable.Table na własne potrzeby
    """

    def __init__(self, parent=None, **kwargs):
        Table.__init__(self, parent, **kwargs)

        self.load_dir = os.getcwd()

        return

    def load(self, files=None):
        # """load from a directory"""
        # if directory == None:
        #     directory = filedialog.askdirectory(parent=self.master,
        #                                               initialdir=os.getcwd(),
        #                                               )
        # if not os.path.exists(directory):
        #     print('directory does not exist')
        #     return
        # if directory:
        #     # prog_bar = Progress(self.master, row=0, column=0, columnspan=2),
        #     excels = []
        #     files = os.listdir(directory)

        #     for file in files:
        #         if file.lower().endswith('.xls') or file.lower().endswith('.xlsx'):
        #             excels.append(os.path.join(directory, file))
        #     plist = PartsListsManager(excels)
        #     df = plist.load_parts_lists()

        #     self.model.df = df
        #     self.adjustColumnWidths()
        #     self.redraw()
        # prog_bar.pb_stop()
        """
        Load multiple files xls to dataframe
        """
        files = filedialog.askopenfilenames(parent=self.master,
                                            initialdir=self.load_dir,
                                            filetypes=[(("Excel", "*.xls*")),
                                                       ("All files", "*.*")])
        if files:
            plist = PartsListsManager(files)
            df = plist.load_parts_lists()
            if df is not None:
                self.model.df = df
                self.adjustColumnWidths()
                self.redraw()

            self.show_log(plist)
            self.load_dir = os.path.basename(files[0])
        return

    def saveAs(self, filename=None):
        """Save dataframe to file"""
        files = [("Excel", "*.xlsx"),
                 ("All files", "*.*")]
        if filename == None:
            filename = filedialog.asksaveasfilename(parent=self.master,
                                                    initialdir=self.currentdir,
                                                    filetypes=files,
                                                    defaultextension=files
                                                    )
        if filename:

            self.filename = filename

            self.model.df.to_excel(filename, index=False)
            self.currentdir = os.path.basename(filename)
        return

    def getSelectedDataFrame(self):
        """Return a sub-dataframe of the selected cells. Will try to convert object
        types to float so that plotting works.

        edited: no convert to float
        """

        df = self.model.df
        rows = self.multiplerowlist
        if not type(rows) is list:
            rows = list(rows)
        if len(rows) < 1 or self.allrows == True:
            rows = list(range(self.rows))
        cols = self.multiplecollist
        # if len(cols) < 1 or self.allcols == True:
        #    cols = list(range(self.cols))
        try:
            data = df.iloc[rows, cols]
        except Exception as e:
            print('error indexing data')
            logging.error("Exception occurred", exc_info=True)
            if 'pandastable.debug' in sys.modules.keys():
                raise e
            else:
                return pd.DataFrame()

        # try to extract numeric
        # colnames = data.columns
        # for c in colnames:
        #     x = pd.to_numeric(data[c], errors='coerce').astype(float)
        #     if x.isnull().all():
        #         continue
        #     data[c] = x
        return data

    def show_log(self, plist):
        """Show the logbox

        Args:
            plist(PartsListsManager): 
        """
        count, log = plist.get_log()
        if count == 1:
            msg = f'Wczytano {count} poprawny plik.\n'
        elif count < 5:
            msg = f'Wczytano {count} poprawne pliki.\n'
        elif count > 4 or count == 0:
            msg = f'Wczytano {count} poprawnych plików.\n'
        if log:
            msg = msg + f'Pliki :\n{log}\nsą nieprawidłowe.'
        messagebox.showinfo(title='Log', message=msg)

    def clear_table(self):
        """Czyści tabele
        """
        model = pd.DataFrame()
        self.model.df = model
        self.redraw()
        return


if __name__ == '__main__':
    app = App()
    # launch the app
    app.mainloop()
