# Gerador de Documentos do Word a partir de Modelos

Este projeto consiste em uma ferramenta que permite a geração automatizada de documentos do Word a partir de modelos predefinidos. É útil em situações em que você deseja criar múltiplas versões de um documento, preenchendo marcadores no modelo com dados de uma planilha Excel.

## Como Funciona

O projeto é composto por duas partes principais:

1. **ExtractDados**: Uma classe que lê dados de uma planilha Excel, normaliza os dados e os disponibiliza para uso.

2. **Geração de Documentos**: O código principal utiliza a biblioteca `docxtpl` para carregar um modelo do Word (documento `.docx`) e substituir marcadores no modelo com os dados da planilha. Os documentos gerados são salvos em um diretório específico.

## Pré-requisitos

- Python
- Biblioteca pandas
- Biblioteca docxtpl

## Uso

1. Configure o ambiente Python e instale as bibliotecas necessárias.
2. Coloque o arquivo de modelo do Word (`.docx`) no diretório apropriado.
3. Defina o nome do modelo e o diretório de saída no código.
4. Execute o código para gerar os documentos a partir dos dados da planilha.

Certifique-se de que os nomes das colunas na planilha Excel correspondam exatamente aos marcadores no modelo do Word para uma substituição precisa.