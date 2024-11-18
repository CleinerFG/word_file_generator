# Gerador de Arquivo Word

## Idiomas da Documentação

- [Leia a documentação em inglês](README.md)

## Visão Geral

Este projeto automatiza o processo de geração de documentos Word com base em dados de uma planilha Excel. A funcionalidade principal da aplicação inclui a leitura de dados de uma planilha Excel especificada, renderização desses dados em um modelo Word predefinido e salvamento dos documentos resultantes em um diretório de saída. Isso permite a geração eficiente de múltiplos documentos com base nas linhas de dados da planilha.

## Classes Principais

### `SheetHandler`

A classe `SheetHandler` é responsável por extrair os dados da planilha Excel.

- **Inicialização**: O construtor aceita o caminho do arquivo da planilha Excel.
- **Métodos de Uso**:
  - `read_sheet()`: Lê a planilha Excel, converte os dados para o formato de dicionário e normaliza os dados.

### `App`

A classe `App` é a classe principal responsável pela geração dos documentos Word a partir dos dados extraídos pela `SheetHandler`.

- **Inicialização**: O construtor aceita os seguintes parâmetros:
  - `outfile`: O nome base para os arquivos de saída.
  - `template_name`: O nome do arquivo modelo Word (padrão é `"template"`).
  - `sheet_name`: O nome do arquivo da planilha Excel (padrão é `"database"`).
- **Métodos de Uso**:
  - `build`: Este é o método principal que orquestra todo o processo. Ele:
    - Cria o diretório de saída.
    - Carrega os dados da planilha Excel.
    - Renderiza cada linha de dados no modelo Word.
    - Salva os documentos renderizados com nomes de arquivos exclusivos.

## Instalação

### 1. **Clone o repositório**:

Clone o repositório com o comando `git clone`:
https://github.com/CleinerFG/word_file_generator.git

### 2. **Instale as dependências**:

Certifique-se de que você tenha o Python 3.x instalado. Você pode instalar as dependências necessárias executando:

`pip install -r requirements.txt`

As dependências necessárias são:

- `numpy`: Complementar para a biblioteca `Pandas`.
- `pandas`: Para leitura e manipulação de arquivos Excel.
- `openpyxl`: Para trabalhar com arquivos Excel nos formatos `.xlsx`.
- `docxtpl`: Para renderizar dados em modelos Word.

## Uso

Para usar a classe `App` e gerar documentos, siga estas etapas:

1. **Prepare a planilha Excel**:

- Certifique-se de que a planilha Excel `database.xlsx` ou outro arquivo `.xlsx` contenha os dados necessários. Cada linha deve representar um conjunto de dados a ser renderizado em um documento. Uma coluna chamada `id` será usada para gerar nomes exclusivos para os documentos.

2. **Prepare o modelo Word**:

- Crie um documento Word modelo `template.docx` ou outro arquivo `.docx` que inclua espaços reservados para os dados. Os espaços reservados devem estar no formato `{{ nome_do_espaco_reservado }}`.

3. **Execute o script**:

- Crie uma instância da classe `App` e chame o método `build()` para gerar os documentos.

### Exemplo

1. **Planilha**: Adicionada à planilha `lista_de_estudantes.xlsx` com as colunas:

- id
- nome
- componente_curricular

  _`Nota`_: A coluna `id` é obrigatória.

2. **Modelo Word**: Adicionado ao documento Word `termo_de_compromisso.docx`. Os espaços reservados no modelo devem ser as colunas da planilha:

- `{{id}}`
- `{{nome}}`
- `{{componente_curricular}}`

3. **Script**:

```python
from classes.app import App

# Crie uma instância da classe App
app = App(outfile="Termo de Compromisso", template_name="termo_de_compromisso", sheet_name="lista_de_estudantes")

# Gere os documentos
app.build()
```
