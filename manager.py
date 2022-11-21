from partslist import PartsList
import pandas as pd


class PartsListsManager:
    def __init__(self, files_path:list):
        self.files_path = files_path
        self.parts_lists = []
        self.df = None

    def load_parts_lists(self):
        for fpath in self.files_path:
            plist = PartsList(fpath)
            plist.load_to_df()

            df = plist.get_df()

            if df is not None:
                self.parts_lists.append(df)

        self.concate_df()
        self.df['ItemID'] = self.df['ItemID'].astype(str)
        return self.df


    def concate_df(self):
        self.df = pd.concat([df for df in self.parts_lists], ignore_index=True)
        # self.df.reset_index(drop=True)
