import pandas as pd
from typing import Dict, List

class SheetHandler:
    """
    Responsible for extracting data from the sheet.
    """
    def __init__(self, dir: str):
        self._dir = dir
        self._data = []
    
    @property
    def dir(self) -> str:
        return self._dir

    @dir.setter
    def dir(self, value: str):
        self._dir = value

    @property
    def data(self) -> List[Dict]:
        return self._data
    
    @data.setter
    def data(self, value: List[Dict]) -> None:
        self._data = value

    def __normalize_data(self) -> List[Dict]:
        normalized_data = []
        for row in self._data:
            dict_row = {}
            for k, v in row.items():
                dict_row[str(k).strip()] = str(v).strip()
            normalized_data.append(dict_row)
        return normalized_data

    def read_sheet(self) -> None:
        try:
            sheet = pd.read_excel(self._dir)
            self.data = sheet.to_dict(orient='records')
            self.data = self.__normalize_data()
        except FileNotFoundError:
            print('The directory for the sheet is incorrect!')

    def printdata(self) -> None:
        if self._data:
            for row in self._data:
                for k,v in row.items():
                    print(f'{k}: {v}', end=' ')
                print('\n')
        else:
            print('No data')
