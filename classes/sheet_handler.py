import pandas as pd
from typing import Dict, List

class SheetHandler:
    """
    Responsible for extracting and processing data from an Excel sheet.
    
    Provides functionality to:
    - Read data from an Excel sheet.
    - Normalize the data by removing leading/trailing spaces from data.
    - Print the data in a user-friendly format.
    """
    
    def __init__(self, dir: str):
        """
        Initializes the SheetHandler with the directory of the Excel sheet.
        
        Args:
            dir (str): The path to the Excel sheet file.
        """
        self._dir = dir
        self._data = []
    
    @property
    def dir(self) -> str:
        """
        Gets the directory of the Excel sheet.
        
        Returns:
            str: The directory path of the Excel sheet.
        """
        return self._dir

    @dir.setter
    def dir(self, value: str):
        """
        Sets a new directory path for the Excel sheet.
        
        Args:
            value (str): The new directory path to set.
        """
        self._dir = value

    @property
    def data(self) -> List[Dict]:
        """
        Gets the data extracted from the sheet.
        
        Returns:
            List[Dict]: The data from the sheet, represented as a list of dictionaries.
        """
        return self._data
    
    @data.setter
    def data(self, value: List[Dict]) -> None:
        """
        Sets the data for the sheet handler.
        
        Args:
            value (List[Dict]): The data to set.
        """
        self._data = value

    def __normalize_data(self) -> List[Dict]:
        """
        Normalizes the data by stripping leading/trailing spaces from keys and values.
        
        Returns:
            List[Dict]: A list of dictionaries with normalized data.
        """
        normalized_data = []
        for row in self._data:
            dict_row = {}
            for k, v in row.items():
                dict_row[str(k).strip()] = str(v).strip()
            normalized_data.append(dict_row)
        return normalized_data

    def read_sheet(self) -> None:
        """
        Reads the Excel sheet from the specified directory and stores the data.
        If the file is found, the data is converted to a list of dictionaries
        and normalized. In case of an invalid path, an error message is printed.
        """
        try:
            sheet = pd.read_excel(self._dir)
            self.data = sheet.to_dict(orient='records')
            self.data = self.__normalize_data()
        except FileNotFoundError:
            print('The directory for the sheet is incorrect!')

    def printdata(self) -> None:
        """
        Prints the extracted data. 
        Each row is printed with key-value pairs. If no data is available, 
        prints 'No data'.
        """
        if self._data:
            for row in self._data:
                for k,v in row.items():
                    print(f'{k}: {v}', end=' ')
                print('\n')
        else:
            print('No data')
