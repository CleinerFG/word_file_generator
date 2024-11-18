import os
from classes.sheet_handler import SheetHandler
from docxtpl import DocxTemplate
from typing import Dict, List

class App:
    """
    Responsible for generating Word documents from data in an Excel sheet.
    
    It reads an Excel sheet, uses a template to generate documents for each row of data,
    and saves them in the output directory.
    """
    
    # Directories for resources and output files
    resource_dir = r'./resource'
    output_dir = r'./output'

    def __init__(self, outfile, template_name="template", sheet_name="database") -> None:
        """
        Initializes with the necessary parameters to generate the output files.
        
        Args:
            outfile (str): The base name for the output files.
            template_name (str, optional): The name of the Word template file (default is "template").
            sheet_name (str, optional): The name of the Excel sheet file (default is "database").
        """
        self._outfile = outfile
        self._template_name = template_name
        self._sheet_name = sheet_name
    
    @property
    def __get_template_dir(self) -> str:
        """
        Returns the full path to the Word template file.
        
        Returns:
            str: The path to the template file.
        """
        return App.resource_dir + f'/{self._template_name}.docx'

    @property
    def __get_sheet_dir(self) -> str:
        """
        Returns the full path to the Excel sheet file.
        
        Returns:
            str: The path to the Excel sheet file.
        """
        return f'{App.resource_dir}/{self._sheet_name}.xlsx'

    @property
    def _get_data(self) -> List[Dict]:
        """
        Reads data from the Excel sheet and returns it as a list of dictionaries.
        
        Returns:
            List[Dict]: A list of dictionaries where each dictionary represents a row in the sheet.
        """
        sheet = SheetHandler(self.__get_sheet_dir)
        sheet.read_sheet()
        return sheet.data

    @classmethod
    def _create_output_dir(cls) -> None:
        """
        Creates the output directory if it doesn't already exist.
        
        This method is called to ensure that the output directory is available before
        saving any generated documents.
        
        Returns:
            None
        """
        if not os.path.exists(cls.output_dir):
            os.makedirs(cls.output_dir)

    def __create_unique_filename(self, identifier) -> str:
        """
        Creates a unique filename for each generated document based on the identifier.
        
        Args:
            identifier (str): The unique identifier to include in the filename (e.g., "id" from the data).
        
        Returns:
            str: The full path to the generated Word document.
        """
        return f'{App.output_dir}/{self._outfile} - {identifier}.docx'

    def build(self):
        """
        Builds the output documents by:
        - Creating the output directory (if not already present).
        - Reading data from the Excel sheet.
        - Generating a Word document for each row of data.
        - Saving the generated documents in the output directory.
        
        Returns:
            None
        """
        App._create_output_dir()
        for row in self._get_data:
            word_file = DocxTemplate(self.__get_template_dir)
            word_file.render(row)
            
            word_file.save(self.__create_unique_filename(row["id"]))
            print(row) 
