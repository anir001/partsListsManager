from manager import PartsListsManager
import os

from tkinter import *
from tkinter import filedialog, messagebox, simpledialog

from pandastable import Table
from pandastable import images
from pandastable.dialogs import addButton

VER = 'v0.1.0'

class App(Frame):
    """Basic frame for the table"""
    def __init__(self, parent=None):
        self.parent = parent
        Frame.__init__(self)
        self.main = self.master
        self.main.geometry('1200x700')
        self.main.title('PartsListsManager ' + VER)


        toolbar = Frame(self.main)
        toolbar.pack(fill='both')

        f = Frame(self.main)
        f.pack(fill='both',expand=True)



        pt = MyTable(f, showtoolbar=False, showstatusbar=True)

        # - toolbar -
        img = images.open_proj()
        addButton(toolbar,
                  'open folder',
                  pt.load, img,
                  side='left',
                  tooltip='open folder PartsLists')
        img = images.save()
        addButton(toolbar,
                  'Load table',
                  pt.saveAs, img,
                  side='left',
                  tooltip='save as xlsx')

        img = images.aggregate()
        addButton(toolbar,
                  'Aggregate',
                  pt.aggregate, img,
                  side='left',
                  tooltip='aggregate')

        img = images.pivot()
        addButton(toolbar,
                  'Pivot',
                  pt.pivot, img,
                  side='left',
                  tooltip='pivot')

        pt.show()
        return

class MyTable(Table):
    def __init__(self, parent=None, **kwargs):
        Table.__init__(self, parent, **kwargs)
        return

    def load(self, directory=None):
        """load from a directory"""
        if directory == None:
            directory = filedialog.askdirectory(parent=self.master,
                                                      initialdir=os.getcwd(),
                                                      )
        if not os.path.exists(directory):
            print('directory does not exist')
            return
        if directory:
            #prog_bar = Progress(self.master, row=0, column=0, columnspan=2),
            excels = []
            files = os.listdir(directory)

            for file in files:
                if file.lower().endswith('.xls') or file.lower().endswith('.xlsx'):
                    excels.append(os.path.join(directory, file))
            plist = PartsListsManager(excels)
            df = plist.load_parts_lists()

            self.model.df = df
            self.adjustColumnWidths()
            self.redraw()
            #prog_bar.pb_stop()
        return

    def saveAs(self, filename=None):
        """Save dataframe to file"""
        files = [(".xlsx", "*.xlsx"),
                   ("All files","*.*")]
        if filename == None:
            filename = filedialog.asksaveasfilename(parent=self.master,
                                                     initialdir = self.currentdir,
                                                     filetypes=files,
                                                     defaultextension = files
                                                     )
        if filename:

            self.filename = filename

            self.model.df.to_excel(filename)
            self.currentdir = os.path.basename(filename)
        return

    # ---- test
    def getSelectedDataFrame(self):
        """Return a sub-dataframe of the selected cells. Will try to convert object
        types to float so that plotting works.

        edited: no convert to float
        """

        df = self.model.df
        rows = self.multiplerowlist
        if not type(rows) is list:
            rows = list(rows)
        if len(rows)<1 or self.allrows == True:
            rows = list(range(self.rows))
        cols = self.multiplecollist
        #if len(cols) < 1 or self.allcols == True:
        #    cols = list(range(self.cols))
        try:
            data = df.iloc[rows,cols]
        except Exception as e:
            print ('error indexing data')
            logging.error("Exception occurred", exc_info=True)
            if 'pandastable.debug' in sys.modules.keys():
                raise e
            else:
                return pd.DataFrame()
        #try to extract numeric
        colnames = data.columns
        # for c in colnames:
        #     x = pd.to_numeric(data[c], errors='coerce').astype(float)
        #     if x.isnull().all():
        #         continue
        #     data[c] = x
        return data




if __name__ == '__main__':
    app = App()
    #launch the app
    app.mainloop()
