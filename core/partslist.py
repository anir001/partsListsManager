import os

import pandas as pd


FIRST_ROWS = ['ItemID',
              'Item Code',
              'Num',
              'source']


class PartsList:
    """Obiekt do przechowywanie pojedyńczego pliku
    """

    def __init__(self, fpath):
        """Funkcja inicjalizacyjna

        Args:
            fpath (str): ścieżka do pliku
        """
        self.df = None
        self.fpath = fpath
        self.is_valid = False

    def get_df(self):
        """Zwraca dataframe

        Returns:
            dataframe: wczytana parts lista
        """
        return self.df

    def is_valid_file(self):
        """Zwraca flagę, czy plik jest poprawnie załadowany

        Returns:
            bool: czy załadowany
        """
        return self.is_valid

    def load_to_df(self):
        """Wczytuje pojedyńczą parts liste z pliku xlsx do dataframe, zaczynając od "row +1"
        """
        try:
            self.df = pd.read_excel(self.fpath)
        except:
            print(f"Nie udało się otworzyć pliku {self.fpath}")
            self.df = None
            return

        if self.df is not None:
            f_row = self._find_first_row()

        if self.is_valid:
            self.df = pd.read_excel(self.fpath, skiprows=f_row + 1)
            self._reindex_df()
            # self.delete_new_line()

        else:
            print(f"Plik {self.fpath}, nie jest poprawną listą.")
            self.df = None

    def _find_first_row(self):
        """Wyszukuje numer wiersza zawierającego str'ItemID' w pierwszej kolumnie pliku excel.
        Wykorzystywane jest to do poprawnego wczytania pliku    

        Returns:
            int: numer wiersza, od którego ma zacząć się wczytywanie do dataframe
        """
        for row in range(self.df.shape[0]):
            if self.df.iloc[row, 0] in FIRST_ROWS:
                self.is_valid = True
                return row

    def _reindex_df(self):
        """Poukładanie dataframe, oraz dołożenie kolumny 'position' - nazwa pliku 
        """
        # Dotyczy list FI
        try:
            col = self.df.pop("Qty")
            self.df.insert(0, col.name, col)
        except:
            pass

        # Dotyczy list EPS
        try:
            self.df.pop("Num")
            self.df.pop("Page")
        except:
            pass

        # dodanie nazwy pliku w pierwszej kolumnie
        _, tail = os.path.split(self.fpath)
        self.df.insert(0, 'source', [
                       tail for i in range(len(self.df.index))])

    def delete_new_line(self):
        """Usunięcie znaków nowej linii
        """
        self.df = self.df.replace('\n', '', regex=True)
