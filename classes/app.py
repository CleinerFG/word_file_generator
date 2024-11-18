import os
from classes.sheet_handler import SheetHandler
from docxtpl import DocxTemplate
from typing import Dict, List

class App:
    resource_dir = r'./resource'
    output_dir = r'./output'

    def __init__(self, outfile, template_name="template", sheet_name="database") -> None:
        self._outfile = outfile
        self._template_name = template_name
        self._sheet_name = sheet_name
    
    @property
    def __get_template_dir(self) -> str:
        return App.resource_dir + f'/{self._template_name}.docx'

    @property
    def __get_sheet_dir(self) -> str:
        return f'{App.resource_dir}/{self._sheet_name}.xlsx'

    @property
    def _get_data(self) -> List[Dict]:
        sheet = SheetHandler(self.__get_sheet_dir)
        sheet.read_sheet()
        return sheet.data

    @classmethod
    def _create_output_dir(cls) -> None:
        if not os.path.exists(cls.output_dir):
            os.makedirs(cls.output_dir)

    def __create_unique_filename(self, identifier):
        return f'{App.output_dir}/{self._outfile} - {identifier}.docx'

    def build(self):
        App._create_output_dir()
        for row in self._get_data:
            word_file = DocxTemplate(self.__get_template_dir)
            word_file.render(row)
            word_file.save(self.__create_unique_filename(row["id"]))
            print(row)