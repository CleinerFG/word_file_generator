# Word File Generator

## Overview

This project automates the process of generating Word documents based on data from an Excel sheet. The core functionality of the application includes reading data from a specified Excel sheet, rendering that data into a predefined Word template, and saving the resulting documents in an output directory. This allows for the efficient generation of multiple documents based on rows of data from the sheet.

## Core Classes

### `SheetHandler`

The `SheetHandler` class is responsible for extracting data from the Excel sheet.

- **Initialization**: The constructor accepts a file path to the Excel sheet.
- **Use Method**:
  - `read_sheet()`: Reads the Excel sheet, converts it to a dictionary format, and normalizes the data.

### `App`

The `App` class is the main class responsible for generating Word documents from the data extracted by `SheetHandler`.

- **Initialization**: The constructor accepts the following parameters:
  - `outfile`: The base name for the output files.
  - `template_name`: The name of the Word template file (defaults to `"template"`).
  - `sheet_name`: The name of the Excel sheet file (defaults to `"database"`).
- **Use Method**:
  - `build`: This is the main method that orchestrates the entire process. It:
    - Creates the output directory.
    - Loads data from the Excel sheet.
    - Renders each row of data into the Word template.
    - Saves the rendered documents with unique filenames.

## Installation

### 1. **Clone the repository**:

Clone the repository with `git clone`.
https://github.com/CleinerFG/word_file_generator.git

### 2. **Install dependencies**:

Ensure that you have Python 3.x installed. You can install the required dependencies by running:

`pip install -r requirements.txt`

The required dependencies are:

- `numpy`: Complementary for the `Pandas` library.
- `pandas`: For reading and handling Excel files.
- `openpyxl`: For working with Excel files in the `.xlsx` formats.
- `docxtpl`: For rendering data into Word templates.

## Usage

To use the `App` class and generate documents, follow these steps:

1. **Prepare the Excel sheet**:
- Ensure that the Excel sheet `database.xlsx` or another `.xlsx` file contains the necessary data. Each row should represent a set of data to be rendered in a document. A column named `id` will be used to generate unique filenames for the documents.

2. **Prepare the Word template**:
- Create a Word document template `template.docx` or another `.docx` file that includes placeholders for the data. The placeholders should be in the form of `{{ placeholder_name }}`.

3. **Run the script**:
- Create an instance of the `App` class and call the `build()` method to generate the documents.

### Example

1. **Sheet**: Added to the spreadsheet `students_list.xlsx` with the columns: 

- id
- name
- curricular_component

    *`Note`*: The id column is mandatory.

2. **Word template**: Added to the word document `commitment_term.docx`. The placeholders in the template must be the sheet collums:

- `{{id}}`
- `{{name}}`
- `{{curricular_component}}`

3. **Script**:
```python
from classes.app import App

# Create an instance of the App class
app = App(outfile="Commitment Term", template_name="commitment_term", sheet_name="students_list")

# Generate the documents
app.build()
