''''
Gerador de arquivos word através de Template

:var diretorio_arq: diretório do arquivo de template
:var nome_arq_save: nome do arquivo a ser salvo

Todos os templates gerados são salvos em arq_modificados
É preciso definir algo para individualização do documento como a matrícula
O nome das colunas do arquivo excel, devem ser o mesmo dos marcadores no docx
    Exemplo:
        plan.xlsx -> colunas: COD, NOME, FUNCAO
        arq.docx -> marcadores: {{COD}}, {{NOME}}, {{FUNCAO}}
'''


from classes.extractdados import ExtractDados
from docxtpl import DocxTemplate

plan = ExtractDados('dados_exemplos.xlsx')
plan.read_sheet()

diretorio_arq = r'termo.docx'
nome_arq_save = 'Termo de entrega de uniformes'

for dados_row in plan.dados:
    arq_word = DocxTemplate(diretorio_arq)
    arq_word.render(dados_row)
    arq_word.save(f'arq_modificados/{nome_arq_save} {dados_row["NOME"]}.docx')
    print(dados_row)