import pandas as pd
import os



class PartsList:
    def __init__(self, fpath):
        self.df = None
        self.founded_item_id = False
        self.fpath = fpath

    def get_df(self):
        return self.df

    def load_to_df(self):
        try:
            self.df = pd.read_excel(self.fpath)
        except:
            self.df = None
            return

        self._find_first_row()

        if self.founded_item_id:
            self._reindex_df()

        else:
            self.df = None

    def _find_first_row(self):
        for i in range(self.df.shape[0]):
            if self.df.iloc[i, 0] == 'ItemID':
                self.df = pd.read_excel(self.fpath, skiprows=i + 1)
                self.founded_item_id = True
                break

    def _reindex_df(self):
        col = self.df.pop("Qty")
        self.df.insert(0, col.name, col)

        # drop last N columns
        # N = 5
        # self.df = self.df.iloc[:, :-N]

        # add file name to first
        head, tail = os.path.split(self.fpath)
        self.df.insert(0, 'position', [tail for i in range(len(self.df.index))])
