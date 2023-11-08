import pandas as pd
from typing import Dict, List

class ExtractDados:
    """
    Classe responsável por extrair os dados da planilha
    """
    def __init__(self, diretorio: str):
        self._diretorio = diretorio
        self._dados = []

    @property
    def diretorio(self) -> str:
        return self._diretorio

    @diretorio.setter
    def diretorio(self, value: str):
        self._diretorio = value

    @property
    def dados(self) -> List[Dict]:
        return self._dados

    def _remove_espacos(self) -> List[Dict]:
        dados_normalizados = []
        for row in self._dados:
            dict_row = {}
            for k, v in row.items():
                dict_row[str(k).strip()] = str(v).strip()
            dados_normalizados.append(dict_row)
        return dados_normalizados

    def read_sheet(self) -> None:
        try:
            sheet = pd.read_excel(self._diretorio)
            self._dados = sheet.to_dict(orient='records')
            self._dados = self._remove_espacos()
        except FileNotFoundError:
            print('O diretório para a planilha está incorreto!')

    def print_dados(self) -> None:
        if self._dados:
            for row in self._dados:
                for k,v in row.items():
                    print(f'{k}: {v}', end=' ')
                print('\n')
        else:
            print('Não há dados')
