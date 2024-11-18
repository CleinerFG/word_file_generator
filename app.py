import os
from classes.extractdados import ExtractDados
from docxtpl import DocxTemplate
from typing import Dict, List

class App:
    template_dir = r'./resource/template.docx'
    output_dir = r'./output'
    def __init__(self, outfile, sheet_name="database") -> None:
        self._outfile = outfile
        self._sheet_name = sheet_name
    
    @property
    def _data(self) -> List[Dict]:
        sheet = ExtractDados(f'./resource/{self._sheet_name}.xlsx')
        sheet.read_sheet()
        return sheet.data
    
    @classmethod
    def _create_output_dir(cls) -> None:
        if not os.path.exists(cls.output_dir):
            os.makedirs(cls.output_dir)

    def build(self):
        App._create_output_dir()
        for row in self._data:
            word_file = DocxTemplate(App.template_dir)
            word_file.render(row)
            word_file.save(f'./output/{self._outfile}-{row["id"]}.docx')
            print(row)