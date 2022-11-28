import os

import pandas as pd

from core.partslist import PartsList


class PartsListsManager:
    """
    Przechowuje wszystkie listy materiałowe w dataframe
    """
    def __init__(self, files_path:list):
        """
        Args:
            files_path (list): lista ścieżek do plików xlsx/xls
        """
        self.files_path = files_path
        self.parts_lists = []
        self.df = None

        self.logs = []
        self.plists_count = 0
        

    def load_parts_lists(self):
        """Ładuje wszystkie pliki do dataframe, następnie łączy w jedno dataframe

        Returns:
            dataframe: złączone dataframe 
        """
        for fpath in self.files_path:
            _, tail = os.path.split(fpath)
            
            plist = PartsList(fpath)
            plist.load_to_df()

            if plist.is_valid_file():
                # self.logger(f'{tail} jest poprawny')
                self.plists_count += 1

                df = plist.get_df()
                self.parts_lists.append(df)

            else:
                self.logger(tail)
                            
        self.concate_df()
        # self.df['ItemID'] = self.df['ItemID'].astype(str)
        
        return self.df
    
    def logger(self, log):
        """Zapisuje wszystkie adresy nieprawidłowych plików

        Args:
            log (str): ścieżka pliku
        """
        self.logs.append(log)

    def get_log(self):
        """Zwraca logi:
            -liczbę wczytanych poprawnych plików
            -liste adresów nieprawidłowych plików
        """
        return self.plists_count, self.logs
        
    def concate_df(self):
        """Łączy wszystkie dataframe w jedno dataframe
        """
        if self.parts_lists:
            self.df = pd.concat([df for df in self.parts_lists], ignore_index=True)
            # self.df.reset_index(drop=True)
